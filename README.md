# hsps Clockify-GSuite integration (v2)

This is a reworking of this original project (https://github.com/viktorbrech/hsps_clockify_gsuite ). It has been substantially changed in the following ways:

* it now aligns with the structure of new the HSPS Clockify instance
* all Python code has been rewritten as Google Apps Script and lives in a single script attached to a Google Sheet. That sheet is thus ready to be copied and customized and used without further dependencies.
* in particular, no communication with a HubSpot portal occurs. No dependency on a HubDB table and no code runs in custom code workflow actions

How to use

* ask Viktor for the Google Sheet, then make a copy of it
* add appropriate values to the "config" sheet (see the explanations on that sheet for guidance)
* add your customers (email domains and HIDs) to the "customers" sheet
* Use the menu "Clockifyiable_Activities" --> "Get project/client/task IDs"
* ensure the functions "log_all_activities" and "refreshSheet_" are run on recurring triggers (this has to be set inside the Google Sheet, check with Viktor to confirm)

More detailed instructions will be added soon.
