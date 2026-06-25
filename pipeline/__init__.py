"""Booking Assistant data pipeline.

Step 1: parse the 2026-2027 booking-request workbooks into a single normalized
booking table, using canonical dictionaries for courses, groups and teachers,
and emit a flags report of everything that needed interpretation or did not
map cleanly.
"""
