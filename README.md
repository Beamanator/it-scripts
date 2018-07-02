# stars-scripts
Different scripts needed for different applications at StARS

Some useful Regular Expressions:
1) `^[0-9]{3}-(CS|AP)[0-9]{8}$|^[0-9]{4}/[0-9]{4}$|^[0-9]{3}-[0-9]{2}C[0-9]{5}$`
   - Useful for UNHCR #s in Google Forms
2)	`^([a-zA-Z0-9_\-\.]+)@stars-egypt\.org(,[ ]{0,1}([a-zA-Z0-9_\-\.]+)@stars-egypt\.org)*$`
    - Multiple Emails (delimiter: ',' or ', ') - email addresses ending with `"...@stars-egypt.org"`
