/*
    How to use:
    1) Copy all of this code into the chrome dev console
    2) run this on the "Advanced Search Results" page
    -> the first result's UNHCR # will be copied to the clipboard.

    How it works:
    1) First, use jQuery to get the UNHCR number from the result table
    2) next, create an 'input' element
    3) put the unhcr number into the input element's value
    4) select the text in the input element (after adding the element to the DOM)
    5) copy the selected text to the clipboard
    6) go back to the previous page (advanced search) to prepare for next client
*/

// step 1
let num = $('table.table.table-striped.grid-table tbody tr td[data-name="HO_REF_NO"')
    .text();

// step 2
let bob = $(document.createElement('input'));
bob.insertAfter('label#pageTitle');

// step 3
bob.val(num);

// step 4
bob[0].select();

// step 5
document.execCommand("Copy");

// step 6
window.history.back();