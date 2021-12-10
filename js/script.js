//set reusable variables 
const reloadBtn =  document.querySelector('.reload');
const formatBtn =  document.querySelector('.format');
const fileSelector = document.getElementById('file-selector');
const reader = new FileReader();
let filename = '';

// detect when file is inserted
fileSelector.addEventListener('change', (event) => {
    const fileList = event.target.files[0];
    readFile(fileList);
});
            
//read file as text which then fires load event to send data to formatting 
const readFile = (file) => { 
    filename = file.name.split('.')[0];
    reader.addEventListener('load', (event) => fileExport(event.target.result));
    reader.readAsText(file);
}

// run file export process when load event triggers
const fileExport = (data) => {
    const workBook = readWorkBook(data);
    const workBookData = formatWorkBookData(workBook);
    const finalData = formatOutput(workBookData[0]);
    saveOutput(finalData);
}

// read the inbound data
const readWorkBook = (workBook) => XLSX.read(workBook, { type: 'binary'});

// format data to json object

const formatWorkBookData = (workBook) => workBook.SheetNames.map((sheetName) => {
            return XLSX.utils.sheet_to_json(workBook.Sheets[sheetName],{
            header:["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB"]
            }
        )
    });


// format data for export
const formatOutput = (data) => {
        // get the values from the data object
        const getKey = Object.values(data[0]);
        // get the keys from the data object
        const setKey = Object.keys(data[0]);
        // setting variables so we can grab them outside of the foreach below
        let email, usage, dmdsid, dmduid, keycode, keycode2, keycode3, keycode4, ins1, ins2, ins3;
        // loop through the object so we can find the fields containg the values we need and retreive the col they are in
        getKey.forEach((key, index) => {
            if (key.toLowerCase().includes('email_address_')) {email = setKey[index]} ;
            if (key.toLowerCase().includes('dmdusage')) {usage= setKey[index]} ;
            if (key.toLowerCase().includes('dmdsid')) {dmdsid = setKey[index]} ;
            if (key.toLowerCase().includes('dmduid')) {dmduid = setKey[index]} ;
            if (key.toLowerCase() === 'keycode') {keycode = setKey[index]} ;
            if (key.toLowerCase() === 'keycode2') {keycode2 = setKey[index]} ;
            if (key.toLowerCase() === 'keycode3') {keycode3 = setKey[index]} ;
            if (key.toLowerCase() === 'keycode4') {keycode4 = setKey[index]} ;
            if (key.toLowerCase() === 'ins1') {ins1 = setKey[index]} ;
            if (key.toLowerCase() === 'ins2') {ins2 = setKey[index]} ;
            if (key.toLowerCase() === 'ins3') {ins3 = setKey[index]} ;
        });
        // take header row out 
        data = data.slice(1);
        // set header for output
        const header = 'EMAIL_ADDRESS_,DMDUSAGE,DMDSID,DMDUID,KEYCODE,KEYCODE2,KEYCODE3,KEYCODE4,INS1,INS2,INS3,HOTMAIL\n';
        // create a new array contaning the newly formatted data we want to export
        return[
            header,
            ...data.map(object => { 
                let setEmail, setUsage, setDmdsid, setDmduid, setKeycode, setKeycode2, setKeycode3, setKeycode4, setIns1, setIns2, setIns3;
                setEmail = (object[email] != undefined) ? `${object[email]}` : ``;
                setUsage = (object[usage] != undefined) ? `${object[usage]}` : ``;
                setDmdsid = (object[dmdsid] != undefined) ? `${object[dmdsid]}` : ``;
                setDmduid = (object[dmduid] != undefined) ? `${object[dmduid]}` : ``;
                setKeycode = (object[keycode] != undefined) ? `${object[keycode]}` : ``;
                setKeycode2 = (object[keycode2] != undefined) ? `${object[keycode2]}` : ``;
                setKeycode3 = (object[keycode3] != undefined) ? `${object[keycode3]}` : ``;
                setKeycode4 = (object[keycode4] != undefined) ? `${object[keycode4]}` : ``;
                setIns1 = (object[ins1]!= undefined) ? `${object[ins1]}` : ``;
                setIns2 = (object[ins2] != undefined) ? `${object[ins2]}` : ``;
                setIns3 = (object[ins3] != undefined) ? `${object[ins3]}` : ``;
                setHotmail = (['@hotmail.ca','@hotmail.co.jp','@hotmail.co.nz','@hotmail.co.uk','@hotmail.co.za','@hotmail.com','@hotmail.com.au','@hotmail.com.br','@hotmail.de','@hotmail.dk','@hotmail.es','@hotmail.fr','@hotmail.it','@hotmail.net','@hotmail.org','@live.ca','@live.co.uk','@live.co.za','@live.com','@live.com.au','@live.de','@live.dk','@live.fr','@live.it','@live.jp','@live.net','@live.nl','@msn.com','@msn.net','@msn.org','@outlook.co.nz','@outlook.com','@outlook.com.au','@outlook.com.br','@outlook.de','@outlook.dk','@outlook.es','@outlook.fr','@outlook.ie','@outlook.in','@outlook.it','@outlook.jp','@outlook.pt','@prodigy.net.mx','@q.com','@windowslive.com'].some(char => setEmail.toLowerCase().endsWith(char))) ? 'Y' : '' ;
                return (`"${setEmail}","${setUsage}","${setDmdsid}","${setDmduid}","${setKeycode}","${setKeycode2}","${setKeycode3}","${setKeycode4}","${setIns1}","${setIns2}","${setIns3}","${setHotmail}"\n`
            )})
            ]
        .join(``);
}
            
// save file to computer
const saveOutput = (data) => { 
        const blob = new Blob([data], { type: "text/plain"});
        const anchor = document.createElement("a");
        const getDateTime = new Date().toLocaleString('en-gb').split(", ");
        const date = getDateTime[0].split("/").reverse().join("");
        const time = getDateTime[1].split(":").join("");
        anchor.download = `${filename}_FLAGGED_${date}${time}.csv`;
        anchor.href = window.URL.createObjectURL(blob);
        anchor.target ="_blank";
        anchor.style.display = "none"; // just to be safe!
        document.body.appendChild(anchor);
        anchor.click();
        document.body.removeChild(anchor);
}

