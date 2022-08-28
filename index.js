const { default: axios } = require("axios");
const xl = require("excel4node");

/*
In this function the axios library is used to do a HTTP request
for the API. Depending on the result of the request, different
messages can be displayed at the console.
*/
async function getData(){
    try{
        const data = axios.get("https://restcountries.com/v3.1/all");
        if (await data){
            console.log("data collected","\nresponse_status_code:",(await data).status,"\n");
            return (await data).data
        } else{
            console.log('Something went wrong!');
        }
    } catch(error){
        console.error("error:", error.code, "\nresponse_status_code:", error.response.status);
    }
}
/*
This functon receives the data collected through the API,
and creates an array on which each object contains data 
about only one country (only the requested data for the
creation of the sheet are put in the array: Name, Capital,
Area and Currencies)
*/
async function buildCountriesArray(receivedData){
    let countriesArray = []
    for (countryData of await(receivedData)){
        let country = {}

        try{country.Name = countryData.name.common}
        catch{country.name = "-"}

        try{country.Capital = countryData.capital.toString()}
        catch{country.capital = "-"}

        try{country.Area = countryData.area.toLocaleString('de-DE',{
            style: "decimal",
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        })}
        catch{country.area = "-"}

        try{country.Currencies = Object.keys(countryData.currencies).toString()}
        catch{country.currencies = "-"}
    
        countriesArray.push(country)
    }
    return countriesArray
}

/*
This function uses the library "excel4node" to create the file 
"CountriesList.xlsx" (which contains the requested sheet). It
receives the array created in the "buildCountriesArray" function 
and iterates through each object of the array (and throught each 
value of each object) and writes them in the file.
*/
async function createSheet(countriesArray){
    var wb = new xl.Workbook();
    var ws = wb.addWorksheet('Sheet 1')

    // style for the title
    var titleStyle = wb.createStyle({
        font: {
          color: '#4F4F4F',
          size: 16,
          bold: true
        },
        alignment:{
            horizontal: 'center',
            vertical: 'center'
        },
    });

    //style for the headers
    var headerStyle = wb.createStyle({
        font: {
          color: '#808080',
          size: 12,
          bold: true
        }
    });

    // write the title in the sheet
    var title = "CountriesList";
    ws.cell(1, 1, 1, 4, true).string(title).style(titleStyle);

    // write the headers in the sheet
    var headers = Object.keys(countriesArray[0]);
    var count = 1;
    for (header of headers){
        ws.cell(2,count).string(header).style(headerStyle);
        count ++;
    }

    // fill the sheet with data 
    rowCount = 3
    colCount = 1

    for(country of countriesArray){
        dataArray = Object.values(country);
        for(data of dataArray){
            ws.cell(rowCount,colCount).string(data)
            colCount++;
        }
        colCount = 1;
        rowCount ++;
    }

    wb.write("CountriesList.xlsx")
}

async function main(){
    const data_colected = getData()
    if(await (data_colected)){
        const countriesArray = buildCountriesArray(data_colected)
        createSheet(await(countriesArray))
    }
}

main()
