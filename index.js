
var Excel = require('exceljs');
const xlsxFile = require('read-excel-file/node');


/*********** Utils - DO NOT CHANGE ***********/

let createObjectList = async (list,fileName) => {
    await xlsxFile('./data/'+fileName+'.xlsx').then((rows) => {
        let colunas = [];
        rows.forEach((col, row_index) => {
            let item = {};
            col.forEach((data, col_index) => { 
                if(row_index == 0){
                    column_name = data.trim();
                    colunas.push(column_name);
                }else{
                    item[colunas[col_index]] = data;
                }
            });
            if(Object.keys(item).length > 0) list.push(item);
        });
    });
    return list;
}

let createNewFile = (list, fileName) => {
    var workbook = new Excel.Workbook();
    var worksheet = workbook.addWorksheet('sheet1');
    var keysTitle = Object.keys(list[0]);
    var objHeading = keysTitle.map(key => {
        return {"header":key,"key":key,width:50}
    });

    worksheet.columns = objHeading;

    list.forEach((item) => {
        worksheet.addRow(item);
    });

    workbook.xlsx.writeFile('./new files/'+fileName+'.xlsx');

    console.log(fileName + ' rows => ', list.length);
};

//////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////

/*********** Declare List of objects you will use ***********/

let products = [];
let collections = [];
let newdata = [];


/*** Add here your processing logic ***/

let processing = () => {

    // Create List of items that contains Title value
    let productNames = products.filter((item) => {
        return item["Title"] != null
    });

    // Populate name
    products.forEach((item) => {
        if(item["Title"] == null){

            let searchedParent = item["Parent"];
            let obj = productNames.find((itemName) => {  
                return (itemName["Parent"] == searchedParent)
            });

            if(obj){
                item["Title"] = obj["Title"];
            }else{
                console.log('*** nao achou nome do produto, ou seja.. nao tem outra row desse produto com title', item);
            }
        }
    });

    // Populate Collection
    products.forEach((item) => {

        let obj = collections.find((itemName) => {  
            return (itemName["Id"] == item["Parent"])
        });

        if(obj){

            // Add collections to product obj
            let smartCollections = (obj["Smart collections"]) ? String(obj["Smart collections"]) : '';
            let customCollections = (obj["Custom collections"]) ? String(obj["Custom collections"]) : '';
            let separador = ( (smartCollections) && (customCollections)) ? ', ' : '';
            let collectionsItem = smartCollections+separador+customCollections;  
            item["Collections"] = collectionsItem;

            // create newdata item ( 1 collection per row)
            let collectionName = collectionsItem.split(',');
            collectionName.forEach((collection) => {
                let colObj = {};
                colObj.Collection = collection.trim();
                let newItem = {...item, ...colObj}
                delete newItem["Collections"];
                newdata.push(newItem);
            });

        }else{
            console.log('nao achou um item correspondente em collections. Ou seja.. o campo Parent do arquivo produtos nao é igual a nenhum campo Id (id de produto) do arquivo collections', item);
        }

    });

}


/*********** Main ***********/
let main = async () => {

    /*** Read Provided Excel files ***/
    await createObjectList(products,'products');
    await createObjectList(collections,'collections');

    /*** Processing logic ***/
    processing();

    /*** Create New Excel Files ***/
    createNewFile(products, Object.keys({products})[0]);
    createNewFile(newdata, Object.keys({newdata})[0]);
    
}

main();