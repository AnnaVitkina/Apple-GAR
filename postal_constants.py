"""Shared postal / bulk-country zone constants for GAR."""

from __future__ import annotations

# All-lanes country zone: integers 0–9 and letters A–Z (BOC26 / RAM bulk country).
BULK_COUNTRY_POSTAL_CODE = ", ".join(
    [str(d) for d in range(10)] + list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
)
