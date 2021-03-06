# it-scripts

Different scripts needed for different it applications

Some useful Regular Expressions:

1. `^[0-9]{3}-(CS|AP)[0-9]{8}$|^[0-9]{4}/[0-9]{4}$|^[0-9]{3}-[0-9]{2}C[0-9]{5}$`
   - Useful for UNHCR #s in Google Forms
2. `^([a-zA-Z0-9_\-\.]+)@gmail\.com(,[ ]{0,1}([a-zA-Z0-9_\-\.]+)@gmail\.com)*$`
   - Multiple Emails (delimiter: ',' or ', ') - email addresses ending with `"...@gmail.com"`

# Other useful things i've learned along the way:

## Google Apps Scripts

1. When getting data from a question (using the `namedValues` prop), the data is returned in an Array, so if you want to check the response text is exactly the same as a string, try using `[0]` to pull the string out.
1. When creating a custom function for sheets, adding a comment block above your function with `@customfunction` can add documentation & auto-complete for that function in the sheet
1. ES6 doesn't usually work (like `let`, `const`, or some useful functions like `map` or concepts like destructuring arrays / objects)
