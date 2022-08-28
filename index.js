const { default: axios } = require("axios");
const xl = require("excel4node");

/*
In this function the axios library is used to do a HTTP request
for the API. Depending on the result of the request, different
messages can be displayed at the console.
*/
async function getData(){
    try{
        const data = await axios.get("https://restcountries.com/v3.1/all");
        if (data){
            console.log("data collected","\nresponse_status_code:", data.status,"\n");
            return data.data
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
function buildCountriesArray(receivedData){
    let countriesArray = []
    for (countryData of receivedData){
        let country = {}

        // Get the name of the country
        if(countryData?.name?.common){
            country.Name = countryData?.name?.common
        }else{
            country.Name = "-"
        }

        // Get the capital of the country
        if(countryData?.capital){
            country.Capital = countryData.capital.toString()
        }else{
            country.Capital = "-"
        }

        // Get the area of the country
        if(countryData?.area){
            country.Area = countryData.area.toLocaleString('pt-BR',{
                style: "decimal",
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            })
        }else{
            country.Area = "-"
        }

        // Get the currencies of the country
        if(countryData?.currencies){
            country.Currencies = Object.keys(countryData.currencies).toString()
        }else{
            country.Currencies = "-"
        }

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
    const wb = new xl.Workbook();
    const ws = wb.addWorksheet('Sheet 1');

    // style for the title
    const titleStyle = wb.createStyle({
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
    const headerStyle = wb.createStyle({
        font: {
          color: '#808080',
          size: 12,
          bold: true
        }
    });

    // write the title in the sheet
    const title = "Countries List";
    ws.cell(1, 1, 1, 4, true).string(title).style(titleStyle);

    // write the headers in the sheet
    const headers = Object.keys(countriesArray[0]);
    let count = 1;
    for (header of headers){
        ws.cell(2,count).string(header).style(headerStyle);
        count ++;
    };

    // fill the sheet with data 
    rowCount = 3
    colCount = 1

    for(country of countriesArray){
        dataArray = Object.values(country);
        for(data of dataArray){
            ws.cell(rowCount,colCount).string(data);
            colCount++;
        }
        colCount = 1;
        rowCount ++;
    }

    wb.write(".xlsx/CountriesList.xlsx")
}

//Here is the main function, from where the application execution begins
async function main(){
    // first, the data is collected via http request
    const dataCollected = await getData()
    if(dataCollected){
        // if the data can be got from the API, the array of countries is built
        const countriesArray = buildCountriesArray(dataCollected)
        // then the sheet is created based on the data contained in the array
        createSheet(countriesArray)
    }
}

main()
