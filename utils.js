const fetch = require("node-fetch");

const saveNonstandardOrderRequest = async(nonstandardOrderRequest) => {

    console.log(`${new Date().toLocaleString('ru')} Gonna save nonstandardOrderRequest: `, nonstandardOrderRequest);

    if('_id' in nonstandardOrderRequest) {
        delete nonstandardOrderRequest._id;
    }
    if('createdAt' in nonstandardOrderRequest) {
        delete nonstandardOrderRequest.createdAt;
    }
    if('updatedAt' in nonstandardOrderRequest) {
        delete nonstandardOrderRequest.updatedAt;
    }
    if('__v' in nonstandardOrderRequest) {
        delete nonstandardOrderRequest.__v;
    }

    try{
        const response = await fetch(`${process.env.BACK_HOST}v2/nonstandardOrderRequests/add`, {
            method: 'post',
            body: JSON.stringify({
                nonstandardOrderRequest,
                st: process.env.SECRET
            }),
            headers: {'Content-Type': 'application/json'}
        });
        const result = await response.json();
        console.log(`${new Date().toLocaleString('ru')} Saving request response: `, result);
        return result;
    } catch(e) {
        console.error(`${new Date().toLocaleString('ru')} Saving request error: `, e);
        return false;
    }
}

const getNonstandardOrderRequestFromDbByTimestamp = async(timestamp) => {
    try{
        console.log(`${new Date().toLocaleString('ru')} Gonna get nonstandard request for date: `, timestamp);
        const response = await fetch(`${process.env.BACK_HOST}v2/nonstandardOrderRequests/timestamp?${new URLSearchParams({ timestamp, st: process.env.SECRET })}`);
        const result = await response.json();
        // console.log(`${new Date().toLocaleString('ru')} Got request result length: `, result.length);
        if(!result.length) {
            return false;
        }
        const lastOne = result.reduce((acc, requestObj) => {
            if(new Date(requestObj.updatedAt) > new Date(acc.updatedAt)) {
                acc = requestObj;
            }
            return acc;
        }, result[0]);
        return lastOne;
    } catch (e) {
        return false;
    }
}

const getLastRowIndexWithValue = (sheet) => {
    console.log(`${new Date().toLocaleString('ru')} Gonna get last row`);
    const newRowCount = sheet.rowCount;
    let lastRowIndexWithValue = 1;
    for(let index = 0; index < newRowCount; index++) {
        console.log(`${new Date().toLocaleString('ru')} Gonna check cell C${index+1}`);
        console.log(`${new Date().toLocaleString('ru')} Value of cell C${index+1}: `, sheet.getCellByA1(`C${index+1}`).value);
        if(!sheet.getCellByA1(`C${index+1}`).value) {
            lastRowIndexWithValue = index;
            break;
        }
    }
    console.log(`${new Date().toLocaleString('ru')} Last row index: `, lastRowIndexWithValue);
    return lastRowIndexWithValue;
}

module.exports = {
    saveNonstandardOrderRequest,
    getLastRowIndexWithValue,
    getNonstandardOrderRequestFromDbByTimestamp
};