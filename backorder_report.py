#!/usr/bin/env python3
"""
backorder_report.py

Reads the most recent "Orders_*.xlsx" export from the Downloads folder
(or a file path passed as an argument), works out which order lines are
currently "backorder" (need to be shipped or re-shipped), and writes
Backorders_<date>.xlsx next to it.

USAGE
    python backorder_report.py
        -> looks in ~/Downloads for the newest file matching Orders_*.xlsx

    python backorder_report.py "C:\\Users\\you\\Downloads\\Orders_202609160217.xlsx"
        -> processes that specific file

LOGIC (see the README block at the bottom of this file for the full
explanation and the assumptions that were made about messy data)
"""

import sys
import re
import glob
import os
from datetime import datetime

import pandas as pd

# --------------------------------------------------------------------------
# CONFIG - tweak these if your data uses different words
# --------------------------------------------------------------------------

# Statuses that mean "hasn't left the building yet"
NOT_YET_SENT_STATUSES = {"New", "Shipment Ready"}

# Statuses that mean "this line is fine, nothing to do"
RESOLVED_STATUSES = {"Delivered", "Sent", "ReShipped"}

# Statuses that mean "something went wrong and it came back / never arrived"
PROBLEM_STATUSES = {"Returned", "Cancelled", "Lost"}

# Short "resend marker" tokens that can appear glued to an order number
# (re, av, uk, pr ...), optionally with a trailing digit like "re2".
SUFFIX_TOKEN_RE = re.compile(r"^(re|av|uk|pr|bo)\d{0,2}$|^\d{0,2}(re|av|uk|pr|bo)$", re.IGNORECASE)

# A bare 1-2 digit token on its own (e.g. the "2" in "re 2 DN73605482",
# meaning "2nd resend") is a resend-count marker, not an order number -
# no real order number in this data is a standalone 1-2 digit number.
BARE_DIGIT_MARKER_RE = re.compile(r"^\d{1,2}$")

# A word counts as part of the order number if it contains a digit.
# Words with no digit at all (re, av, replacement, gift, event, test ...)
# are treated as notes/markers, not as part of the order identity.


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def find_latest_orders_file(downloads_dir):
    candidates = glob.glob(os.path.join(downloads_dir, "Orders_*.xlsx"))
    candidates = [c for c in candidates if "Backorders_" not in os.path.basename(c)]
    if not candidates:
        return None
    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates[0]


def tokenize_order_no(raw):
    """Split an OrderNo cell into (order_number_tokens, other_tokens)."""
    if pd.isna(raw):
        return [], []
    text = str(raw).strip()
    tokens = re.split(r"[\s+,]+", text)
    order_tokens = []
    other_tokens = []
    for tok in tokens:
        tok = tok.strip()
        if not tok:
            continue
        if SUFFIX_TOKEN_RE.match(tok) or BARE_DIGIT_MARKER_RE.match(tok):
            other_tokens.append(tok.lower())
        elif re.search(r"\d", tok):
            order_tokens.append(tok.upper())
        else:
            other_tokens.append(tok.lower())
    return order_tokens, other_tokens


def build_order_key(raw):
    order_tokens, _ = tokenize_order_no(raw)
    if order_tokens:
        return "+".join(sorted(set(order_tokens)))
    # Fallback: nothing that looked like an order number (e.g. "test",
    # "Gift NSS Paris event") -> just use the normalised raw text so it
    # doesn't accidentally collide with anything else.
    if pd.isna(raw):
        return ""
    return re.sub(r"\s+", " ", str(raw).strip().upper())


class _UnionFind:
    """Tiny union-find so that two orders which start out separate but
    later get combined into one shipment (e.g. 'av DN123 + DN456') are
    recognised as the SAME order history, not two unrelated ones."""

    def __init__(self):
        self.parent = {}

    def find(self, x):
        self.parent.setdefault(x, x)
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            # deterministic winner so the key is stable regardless of
            # processing order
            if ra < rb:
                self.parent[rb] = ra
            else:
                self.parent[ra] = rb


def build_order_keys(order_no_series):
    """Return an OrderKey per row. Order-number tokens are linked with a
    union-find across the WHOLE file first: if DN123 ever appears together
    with DN456 in any single OrderNo cell (e.g. a later resend combines two
    previously separate orders into one shipment), every row that mentions
    EITHER number is grouped into the same order history."""
    per_row_tokens = [tokenize_order_no(raw)[0] for raw in order_no_series]

    uf = _UnionFind()
    for tokens in per_row_tokens:
        for t in tokens:
            uf.find(t)  # register the node
        for t in tokens[1:]:
            uf.union(tokens[0], t)

    keys = []
    for raw, tokens in zip(order_no_series, per_row_tokens):
        if tokens:
            keys.append(uf.find(tokens[0]))
        elif pd.isna(raw):
            keys.append("")
        else:
            keys.append(re.sub(r"\s+", " ", str(raw).strip().upper()))
    return keys


REFUND_RE = re.compile(r"efund|erstatt", re.IGNORECASE)
# "efund" catches "refund" and the "fefund" typo.
# "erstatt" is the shared stem of the German word for refund and its
# forms - Rückerstattung, Rückerstattungen, rückerstattet, Erstattung,
# Ruckerstattung (no umlaut) all contain it, regardless of prefix.

# "AV" (address verification) marker: catches standalone "AV" and the
# spelled-out phrase. Word-boundary on AV so it doesn't fire on plain
# street-address text like "Av. Zaragoza, 24".
AV_RE = re.compile(r"\bAV\b|address\s+verification", re.IGNORECASE)


def has_refund_note(comment, notes):
    text = " ".join(str(x) for x in (comment, notes) if pd.notna(x))
    return bool(REFUND_RE.search(text))


def has_av_note(comment, notes):
    text = " ".join(str(x) for x in (comment, notes) if pd.notna(x))
    return bool(AV_RE.search(text))


# --------------------------------------------------------------------------
# Main processing
# --------------------------------------------------------------------------

def process(path):
    df = pd.read_excel(
        path,
        dtype={"OrderNo": str, "Receipt No": str, "SKU": str, "Customer Code": str},
    )

    # Drop fully blank rows (spacer rows some of these exports contain)
    df = df[df["OrderNo"].notna()].copy()
    df.reset_index(drop=True, inplace=True)

    # Defensive: strip stray whitespace from Status so it always matches
    # the sets above (past exports have had trailing spaces on other text
    # columns, e.g. order numbers).
    df["Status"] = df["Status"].astype(str).str.strip()

    # --- identity columns -------------------------------------------------
    df["OrderKey"] = build_order_keys(df["OrderNo"])
    df["AttemptText"] = df["OrderNo"].str.strip()

    # An "attempt" = one exact OrderNo string within one OrderKey.
    # Its date = earliest 'Order upload date' seen for that exact text.
    attempt_dates = (
        df.groupby(["OrderKey", "AttemptText"])["Order upload date"]
        .min()
        .reset_index()
        .rename(columns={"Order upload date": "AttemptDate"})
    )

    # Rank attempts chronologically within each OrderKey -> xSent
    attempt_dates.sort_values(["OrderKey", "AttemptDate"], inplace=True)
    attempt_dates["xSent"] = attempt_dates.groupby("OrderKey").cumcount() + 1

    # Latest attempt per OrderKey
    idx_latest = attempt_dates.groupby("OrderKey")["AttemptDate"].idxmax()
    latest_attempt_text = attempt_dates.loc[idx_latest, ["OrderKey", "AttemptText"]]
    latest_attempt_text = latest_attempt_text.rename(columns={"AttemptText": "LatestAttemptText"})

    df = df.merge(attempt_dates[["OrderKey", "AttemptText", "xSent"]], on=["OrderKey", "AttemptText"], how="left")
    df = df.merge(latest_attempt_text, on="OrderKey", how="left")
    df["IsLatestAttempt"] = df["AttemptText"] == df["LatestAttemptText"]

    # --- backorder decision -------------------------------------------------
    def decide(row):
        status = row["Status"]
        if not row["IsLatestAttempt"]:
            return "N", "Superseded - sent again later"
        if status in NOT_YET_SENT_STATUSES:
            return "Y", "Not yet sent"
        if status in RESOLVED_STATUSES:
            return "N", "Sent / delivered"
        if status in PROBLEM_STATUSES:
            if has_refund_note(row.get("Comment"), row.get("Notes")):
                return "N", "Refunded, not a backorder"
            if status == "Cancelled" and has_av_note(row.get("Comment"), row.get("Notes")):
                return "Y", "Cancelled - AV/address verification needed, not resent"
            return "Y", f"{status} and not sent again"
        # Cancelled orders sometimes carry no formal status at all - if
        # the row is still awaiting address verification and was never
        # resent, treat it as a backorder too.
        if has_av_note(row.get("Comment"), row.get("Notes")) and not has_refund_note(row.get("Comment"), row.get("Notes")):
            return "Y", f"AV/address verification needed (status: {status!r})"
        return "N", f"Unrecognised status: {status!r} - check manually"

    decisions = df.apply(decide, axis=1, result_type="expand")
    df["Backorder"] = decisions[0]
    df["BackorderReason"] = decisions[1]

    # Clean up helper columns the user didn't ask for but keep the useful ones
    df.drop(columns=["LatestAttemptText", "IsLatestAttempt"], inplace=True)

    return df


def main():
    if len(sys.argv) > 1:
        src = sys.argv[1]
    else:
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        src = find_latest_orders_file(downloads)
        if src is None:
            print(f"No Orders_*.xlsx file found in {downloads}. "
                  f"Pass a file path as an argument instead.")
            sys.exit(1)
        print(f"Using newest export found: {src}")

    df = process(src)

    out_dir = os.path.dirname(os.path.abspath(src))
    out_name = f"Backorders_{datetime.now().strftime('%Y%m%d')}.xlsx"
    out_path = os.path.join(out_dir, out_name)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Orders", index=False)

    n_y = (df["Backorder"] == "Y").sum()
    print(f"Done. {n_y} backorder line(s) out of {len(df)} total lines.")
    print(f"Saved: {out_path}")


if __name__ == "__main__":
    main()


# --------------------------------------------------------------------------
# README - read this if a result looks wrong
# --------------------------------------------------------------------------
#
# WHAT COUNTS AS AN "ORDER"
#   Order numbers are grouped together (an "OrderKey") by pulling out every
#   token in the OrderNo cell that contains a digit (e.g. DN77750994,
#   R26G80069, 10688). Tokens with NO digit at all (re, av, uk, pr,
#   "replacement", "gift", "test", ...) are treated as markers/notes, not
#   as part of the order's identity, so "DN123 re" and "DN123" are
#   recognised as the SAME order, at two different points in time (two
#   "attempts"). A bare 1-2 digit token on its own (e.g. "re 2 DN123",
#   meaning "2nd resend") is treated the same way - it's a resend
#   counter, not a separate order number.
#
#   Grouping is transitive across the WHOLE file, not just within one
#   cell: if order DN123 was ever combined with DN456 in a single OrderNo
#   cell (e.g. two separately cancelled orders get shipped together later
#   as "av DN123 + DN456"), every row mentioning EITHER number - including
#   their original, separate, pre-combination attempts - is linked into
#   one order history. This is what makes "two orders were cancelled,
#   then combined and sent together" resolve correctly instead of leaving
#   the two original cancelled lines stranded as false backorders.
#
# WHAT COUNTS AS A "SEND ATTEMPT" (xSent)
#   Every distinct exact OrderNo TEXT within one OrderKey is one attempt
#   (checked against the data: the same exact OrderNo text always shares
#   the same order-upload date, so this is a safe way to split attempts).
#   Attempts are numbered 1, 2, 3... in the order they were uploaded.
#
# WHAT COUNTS AS "BACKORDER"
#   Only the MOST RECENT attempt of an order can be a backorder - if an
#   order was shipped, came back, and was shipped again (a later attempt
#   exists), the earlier attempt is marked N ("superseded") and only the
#   newest attempt is judged on its own status:
#     - Status is New / Shipment Ready           -> Backorder = Y (not sent yet)
#     - Status is Delivered / Sent / ReShipped    -> Backorder = N (handled)
#     - Status is Returned / Cancelled / Lost     -> Backorder = Y, UNLESS the
#       Comment or Notes column contains the word "refund" (also catches
#       the "fefund" typo) or the German "Rückerstattung"/"erstattet"/
#       "Erstattung" (and umlaut-less "Ruckerstattung"), in which case
#       it's N (refunded, nothing to ship).
#     - Cancelled orders whose Comment/Notes mention "AV" or "address
#       verification" (and aren't a refund) get a dedicated reason label
#       ("Cancelled - AV/address verification needed, not resent") so
#       they're easy to filter for, but they were already Y under the
#       plain "Cancelled" rule above - this is mainly there to make sure
#       they stay Y even if Status is blank/unusual, see below.
#     - Any status OUTSIDE the three buckets above, if it still mentions
#       "AV"/"address verification" (and isn't a refund) -> Backorder = Y
#       too, so an address-verification case never slips through just
#       because its Status field is missing or unexpected.
#     - Any other/unexpected status               -> Backorder = N, but the
#       BackorderReason column says "Unrecognised status" so you can check
#       it by hand - this should not normally happen.
#
#   Line items are decided individually (not the whole order at once),
#   because a small number of orders (~1% in the sample file) have some
#   items delivered and others cancelled within the very same shipment.
#
# ASSUMPTIONS / KNOWN LIMITATIONS - worth a manual spot-check
#   1. "Not yet sent" (New / Shipment Ready) is flagged as a backorder
#      immediately, even if the order was only placed minutes ago. If you
#      only want orders stuck for a while, filter Backorders Only further
#      by "Order upload date".
#   2. Refund detection only looks for the text "refund" (or "fefund").
#      Refunds that are recorded differently won't be caught - double
#      check any BackorderReason = "Returned/Cancelled/Lost and not sent
#      again" cases that you know were actually refunded.
#   3. A handful of OrderNo values are messy free text with no digits at
#      all (e.g. "test", "Gift NSS Paris event") - these can't be matched
#      to any other row, so they're always treated as a single, final
#      attempt (xSent = 1).
#   4. "Cancelled" is treated the same as "Returned"/"Lost" (needs
#      resending unless refunded) per your instructions. If some
#      cancellations should never be treated as backorders, let me know
#      and I'll add another rule.
