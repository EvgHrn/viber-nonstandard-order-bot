const fetch = require("node-fetch");

export const saveNonstandardOrderRequest = async(nonstandardOrderRequest) => {

    console.log(`${new Date().toLocaleString('ru')} Gonna save nonstandardOrderRequest: `, nonstandardOrderRequest);

    const response = await fetch(`${process.env.BACK_HOST}v2/nonstandardOrderRequests/add`, {
        method: 'post',
        body: JSON.stringify({
            nonstandardOrderRequest,
            st: process.env.SECRET
        }),
        headers: {'Content-Type': 'application/json'}
    });
    const result = await response.json();
    console.log(`${new Date().toLocaleString('ru')} Saving request result: `, result);
}

export const getLastRowIndexWithValue = (sheet) => {
    const newRowCount = sheet.rowCount;
    let lastRowIndexWithValue = 0;
    for(let index = 0; index < newRowCount; index++) {
        console.log(`${new Date().toLocaleString('ru')} Gonna check cell C${index+1}`);
        console.log(`${new Date().toLocaleString('ru')} Value of cell C${index+1}: `, sheet.getCellByA1(`C${index+1}`).value);
        if(!sheet.getCellByA1(`C${index+1}`).value) {
            lastRowIndexWithValue = index;
            break;
        }
    }
}