

' Place the following in the Code Builder within the form in the "Before Update" section. Where xxx is the ID column name.
Call AuditChanges("xxx")

' Each field you want to log must have the word "Audit" in the tag field.
