var output = '';

// loop through table on rips.247lib.com/StarsSup/User/UserDetails
$('table.webGrid tbody tr').each(function(index, elem) {
    let rowArr = elem.getElementsByTagName('td');
    
    // get user data from table
    let userID = rowArr[0]
        .getElementsByTagName('a')[0].innerText;
    let username = rowArr[1].innerText;
    let password = rowArr[2].innerText;
    // let accessLevel = rowArr[3];
    let teamID = rowArr[4].innerText;
    // let reportGroup = rowArr[5];
    let email = rowArr[6].innerText;

    // if password has "St@RS" in it, it probably hasn't been changed
    if ( password.indexOf('St@RS') !== -1 )
        output += userID + "\t" +
            username + "\t" +
            password + "\t" +
            teamID + "\t" + 
            email + "\t" +
            "\n";
});

console.log(output);