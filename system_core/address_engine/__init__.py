"""How an address names its settlement, street and house.

These modules come from Audion Address Processor (system_core/address_engine):
audion_address_core.py, address_slots.py and models.py are exact copies. DocFlow
only needs the naming - г., п., ул., д., к. - so the registry lookup, postal codes
and districts of that engine stay there, and this package imports nothing else.

Change the files in Address Processor and copy them here unchanged;
tests/test_audit_rule_files.py compares the copies when both projects sit side by side.
"""
