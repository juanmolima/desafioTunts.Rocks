const axios = require('axios');
const xl = require('excel4node');


// get the data from the api
getData = async ( ) => { 
    const response = await axios.get("https://restcountries.com/v3.1/all")
    const data = response.data
    return data
}

  
getData().then(data => {

// format the data to be used in the excel file
const countryData = {
    name: "",
    capital: "",
    area: "",
    currency: ""
};

countries = [];

data.forEach(e => { 
    countryData.name = e.name.common;

    e.capital ? countryData.capital = e.capital : countryData.capital = "-";
    e.area ? countryData.area = e.area : countryData.area = "-";
    countryData.currency =  Object.keys(
        e.currencies || { "-": null }
      ).join(",");
    countries.push({...countryData});
    countries.sort((a, b) => a.name.localeCompare(b.name));
})
    

countries.push(countryData);

// create the excel file
const wb = new xl.Workbook();
const ws = wb.addWorksheet('CountriesList');

// create the styles for the excel file

const styleNumber = wb.createStyle({
    numberFormat: "#,##0.00",
  });

const styleTitle = wb.createStyle({ 
    font: {
        size: 16,
        bold: true,
        color: '#4F4F4F'
    },
    alignment: { 
        horizontal: 'center',
    }
});

const styleColumns = wb.createStyle({
    font: {
        size: 12,
        bold: true,
        color: '#808080'
    }    
 });

// create the header of the excel file
ws.cell(1, 1, 1, 4, true)
.string('Countries List')
.style(styleTitle);

// create the columns of the excel file
let columnHeading = ["Name", "Capital", "Area", "Currencies"];

let columnHeadingIndex = 1;
columnHeading.forEach(heading => {
    ws.cell(2, columnHeadingIndex++)
    .string(heading)
    .style(styleColumns);
    
});

// create the rows of the excel file
let rowIndex = 3;
countries.forEach((record) => {
  let columnIndex = 1;
  Object.keys(record).forEach((columnName) => {
    if (columnIndex === 3) {
      ws.cell(rowIndex, columnIndex++)
        .number(record[columnName])
        .style(styleNumber);
    } else {
      ws.cell(rowIndex, columnIndex++).string(record[columnName]);
    }
  });
  rowIndex++;
});
  ws.column(3).setWidth(13);

  wb.write("CountriesList.xlsx"); 
})
.catch(error => { console.log(error) } );

getData()

