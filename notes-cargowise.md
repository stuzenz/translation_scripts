Goal: Mapping UI Fields to XML Tags and Database Fields

Given this context, the goal is to create a mapping tool that links CargoWise UI fields to their corresponding XML tags and underlying database fields. In other words, for each field a user sees on the CargoWise screen, we want to know:

    Which XML element or attribute represents that field in CargoWise’s import/export messages.

    Which database table and column store that field’s value in the CargoWise database (accessible via a read-only connection).

Achieving this mapping will enable you (or developers) to more easily integrate and verify data – for example, ensuring that an XML message is populating the correct field in CargoWise, or vice versa, by tracing it to the DB field.

This is a non-trivial task because the naming of fields can differ across the UI, XML, and DB:

    The UI labels are user-friendly names (e.g. "House Bill", "Local Client").

    The internal field identifiers in the application and database are typically coded (e.g. JH_JobNum or OH_FullName).

    The XML tags have their own naming conventions (often similar to the internal names but without table prefixes).

Fortunately, CargoWise provides ways to discover these mappings, and with some scripting and use of documentation, we can automate much of the work. Below, we outline an approach to build a mapping tool, leveraging CargoWise UI shortcuts, official documentation/XSDs, and Python for automation.

CargoWise UI – Field Identifier: CargoWise One has a built-in shortcut to reveal the internal field name behind a UI field. By focusing on a field in the UI and pressing Ctrl + Shift + R, the system will display the actual data field name for that field

. This is extremely useful – it gives you the code name which usually corresponds directly to the database column. For example, if you place the cursor in an organization name field and press the shortcut, it might show OH_FullName (where "OH" is a prefix and "FullName" is the field)

. In this case:

    UI Field Label: Full Name (on an Organization screen)

    Internal Name/DB Field: OH_FullName

    XML Tag: <FullName> in the Org XML schema (the XML uses the same field name without the two-letter prefix
    
    ).