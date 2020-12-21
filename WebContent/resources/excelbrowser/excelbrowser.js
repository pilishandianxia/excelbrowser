var ExcelBrowser = function( ){

    var sheetNamesObject = null;
    var sheetsObject = null;
    var previewSheetsObject = null;
    var editingSheet = null;
        
    ExcelBrowser.onExcelFileSubmit = function(){

    };

    ExcelBrowser.excelFileUploading = function(){
        
    };

    ExcelBrowser.excelFileUploaded = function(json){

        sheetNamesObject = eval('(' + json + ')');

        createSheetNameNoteTable();

    };

    ExcelBrowser.listAllSheets = function(){

        sheetsObject = jqueryAjax("", "listAllSheets", "");

        createSheetsTable();
        
    };

    ExcelBrowser.delSheets = function(){

        if(sheetsObject == null){

            return;

        }

        var sheetListObject = {};

        var sheetIds = [];

        sheetListObject.sheetIds = sheetIds;
        
        for(var i = 0; i < sheetsObject.sheets.length; i++){

            var sheet = sheetsObject.sheets[i];

            var input = document.getElementById("delSheet" + sheet.id);
            if(input.checked == true){

                sheetIds.push( sheet.id );
                
            }

        }

        var jsonString = JSON.stringify( sheetListObject );

        jqueryAjax("sheetIds=" + jsonString, "delSheets", "");
        
        
    };

    ExcelBrowser.previewSheets = function(){

        if(sheetsObject == null){

            return;

        }

        var sheetListObject = {};

        var sheetIds = [];

        sheetListObject.sheetIds = sheetIds;
        
        for(var i = 0; i < sheetsObject.sheets.length; i++){

            var sheet = sheetsObject.sheets[i];

            var input = document.getElementById("delSheet" + sheet.id);
            if(input.checked == true){

                sheetIds.push( sheet.id );
                
            }

        }

        var jsonString = JSON.stringify( sheetListObject );

        previewSheetsObject = jqueryAjax("sheetIds=" + jsonString, "previewSheets", "");

        createPreviewSheetsTable();
        
        
    };
    
    ExcelBrowser.downloadSheets = function(){

        if(sheetsObject == null){

            return;

        }

        //var sheetListObject = {};

        //var sheetIds = [];

        //sheetListObject.sheetIds = sheetIds;

        var parameters = "";
        for(var i = 0; i < sheetsObject.sheets.length; i++){

            var sheet = sheetsObject.sheets[i];

            var input = document.getElementById("delSheet" + sheet.id);
            if(input.checked == true){

                if( i == 0 ){

                    parameters += "sheetId0=" + sheet.id;
                }
                else{

                    parameters += "&sheetId" + i + "=" + sheet.id;

                }
                //sheetIds.push( sheet.id );
                
            }

        }

        //var jsonString = JSON.stringify( sheetListObject );

        window.location = "downloadSheets?" + parameters;
        //jqueryAjax("", "downloadSheets?sheetIds=" + jsonString, "");
        
        
    };

    ExcelBrowser.editSheet = function(){

        if( editingSheet == null ){

            return;
            
        }
        
        var alias = document.getElementById( "editSheetAlias" ).value;
        var description = document.getElementById( "editSheetDescription" ).value;

        var jsonObject = {};

        jsonObject.sheetId = editingSheet.id;
        jsonObject.alias = alias;
        jsonObject.description = description;

        var jsonString = JSON.stringify( jsonObject );

        jqueryAjax("sheetNew=" + jsonString, "editSheet", "");
        
        
    };

    ExcelBrowser.searchSheets = function(){

        var alias = document.getElementById( "searchByAlias" ).value;
        var description = document.getElementById( "searchByDescription" ).value;

        var searchSheetObject = {};

        searchSheetObject.alias = alias;
        searchSheetObject.description = description;

        var jsonString = JSON.stringify( searchSheetObject );

        sheetsObject = jqueryAjax("searchCondition=" + jsonString, "searchSheets", "");

        createSheetsTable();
        

    };

    ExcelBrowser.pagingSheets = function(){

        var sheetCountObject = jqueryAjax("", "sheetCount", "");

        var sheetCount = sheetCountObject.sheetCount;

        var pageNumber = document.getElementById( "pageNumber" ).value;
        pageNumber = new Number( pageNumber ).valueOf();

        var sheetPerPage = document.getElementById( "sheetPerPage" ).value;
        sheetPerPage = new Number( sheetPerPage ).valueOf();

        var pagingObject = {};
        
        pagingObject.pageNumber = pageNumber;
        pagingObject.sheetPerPage = sheetPerPage;

        var jsonString = JSON.stringify( pagingObject );
        
        sheetsObject = jqueryAjax("paging=" + jsonString, "pagingSheets", "");

        createSheetsTable();
        

    };
    
    function jqueryAjax(params, url, async){
        
    	var returnData = null;
        
    	$.ajax({
    		type: "POST",
    		async: async,
    		url: url,
    		data: params,
    		dataType: "json",
    		success: function( data ){
    		
    			returnData = data;
                
    		}
            
    	});
        
    	return returnData;
    }

    function submitSheetNameNotes(){

        var jsonObject = {};

        jsonObject.excelFileId = sheetNamesObject.excelFileId;

        jsonObject.sheetNameNotes = [];
        
        for( var i = 0; i < sheetNamesObject.sheetNames.length; i++ ){

            var sheetNameNoteObject = {};

            sheetNameNoteObject.sheetName = sheetNamesObject.sheetNames[i];
            sheetNameNoteObject.alias = document.getElementById( "alias" + i ).value;
            sheetNameNoteObject.description = document.getElementById( "description" + i ).value;

            jsonObject.sheetNameNotes.push(sheetNameNoteObject);

        }

        var jsonString = JSON.stringify(jsonObject);

        
        jqueryAjax("sheetNameNotes=" + jsonString, "excelToMysql", "");
        
            
    }

    function createSheetsTable(){

        var table = document.getElementById( "sheetsTable" );

        var tr = document.createElement( "tr" );
        table.appendChild( tr );

        var td = document.createElement( "td" );
        td.innerText = "Del";            
        tr.appendChild( td );

        var td = document.createElement( "td" );
        td.innerText = "Edit";            
        tr.appendChild( td );

        var td = document.createElement( "td" );
        td.innerText = "Id";            
        tr.appendChild( td );
        
        var td = document.createElement( "td" );
        td.innerText = "Table name";            
        tr.appendChild( td );
        
        var td = document.createElement( "td" );
        td.innerText = "Sheet name";            
        tr.appendChild( td );
        
        var td = document.createElement( "td" );
        td.innerText = "Alias";            
        tr.appendChild( td );
        
        var td = document.createElement( "td" );
        td.innerText = "Description";            
        tr.appendChild( td );
        
        var td = document.createElement( "td" );
        td.innerText = "Department";            
        tr.appendChild( td );
        
        var td = document.createElement( "td" );
        td.innerText = "Records";            
        tr.appendChild( td );
        
        var td = document.createElement( "td" );
        td.innerText = "Time";            
        tr.appendChild( td );

        for(var i = 0; i < sheetsObject.sheets.length; i++){

            var sheet = sheetsObject.sheets[i];
            
            var tr = document.createElement( "tr" );
            table.appendChild( tr );

            var td = document.createElement( "td" );
            tr.appendChild( td );
            var input = document.createElement( "input" );
            input.id = "delSheet" + sheet.id;
            input.type = "checkbox";
            input.name = "delSheet";
            td.appendChild( input );

            var td = document.createElement( "td" );
            tr.appendChild( td );
            var input = document.createElement( "input" );
            input.id = "editSheet" + sheet.id;
            input.type = "button";
            input.value = "Edit";
            input.onclick = function(){

                editSheet( this.id );

            };            
            td.appendChild( input );
            
            
            var td = document.createElement( "td" );
            td.innerText = sheet.id;            
            tr.appendChild( td );
            
            var td = document.createElement( "td" );
            td.innerText = sheet.tableName;            
            tr.appendChild( td );
            
            var td = document.createElement( "td" );
            td.innerText = sheet.sheetName;            
            tr.appendChild( td );
            
            var td = document.createElement( "td" );
            td.innerText = sheet.alias;            
            tr.appendChild( td );
            
            var td = document.createElement( "td" );
            td.innerText = sheet.description;            
            tr.appendChild( td );
            
            var td = document.createElement( "td" );
            td.innerText = sheet.department;            
            tr.appendChild( td );
            
            var td = document.createElement( "td" );
            td.innerText = sheet.records;            
            tr.appendChild( td );

            var dateTime = new Date();
            dateTime.setTime(sheet.timestamp);
            
            var td = document.createElement( "td" );
            td.innerText = dateTime.toString();            
            tr.appendChild( td );
            
            
        }
        
    }
    
    function createSheetNameNoteTable(){

        var table = document.getElementById( "sheetNameNotesTable" );

        for( var i = 0; i < sheetNamesObject.sheetNames.length; i++ ){

            //First line, for sheet name
            var tr = document.createElement( "tr" );
            table.appendChild( tr );

            var td = document.createElement( "td" );
            td.innerText = sheetNamesObject.sheetNames[i];            
            tr.appendChild( td );

            var td = document.createElement( "td" );
            tr.appendChild( td );

            var td = document.createElement( "td" );
            tr.appendChild( td );

            //Second line, for sheet's alias
            var tr = document.createElement( "tr" );
            table.appendChild( tr );
            
            var td = document.createElement( "td" );
            tr.appendChild( td );

            var td = document.createElement( "td" );
            td.innerText = "Alias";
            tr.appendChild( td );
            
            var td = document.createElement( "td" );
            tr.appendChild( td );
            var input = document.createElement( "input" );
            input.id = "alias" + i;
            input.type = "text";
            input.size = "40";
            td.appendChild(input);


            //Third line, for sheet's description
            var tr = document.createElement( "tr" );
            table.appendChild( tr );
            
            var td = document.createElement( "td" );
            tr.appendChild( td );

            var td = document.createElement( "td" );
            td.innerText = "Description";
            tr.appendChild( td );
            
            var td = document.createElement( "td" );
            tr.appendChild( td );
            var input = document.createElement( "input" );
            input.id = "description" + i;
            input.type = "text";
            input.size = "40";
            td.appendChild(input);                        


        }

        //Last line, for submint button
        var tr = document.createElement( "tr" );
        table.appendChild( tr );

        var td = document.createElement( "td" );
        td.colSpan = "3";
        tr.appendChild( td );

        var input = document.createElement( "input" );
        input.type = "button";
        input.value = "Submit";
        input.onclick = submitSheetNameNotes;
        
        td.appendChild(input);

        
    }

    function createPreviewSheetsTable(){

        var sheets = previewSheetsObject.sheets;

        for(var i = 0; i < sheets.length; i++ ){

            createPreviewSheetTable( sheets[i] );

        }

    }

    function createPreviewSheetTable( sheet ){

        createSheetInformationTable( sheet.sheetinformation );

        createRecordsTable( sheet.columninformation, sheet.records );

    }

    function createSheetInformationTable( sheetinformation ){

        var previewSheetsTable = document.getElementById( "previewSheetsTable" );

        var tr = document.createElement( "tr" );
        previewSheetsTable.appendChild( tr );

        var td = document.createElement( "td" );
        tr.appendChild( td );

        var table = document.createElement( "table" );
        td.appendChild( table );
        
        var tr = document.createElement( "tr" );
        table.appendChild( tr );
       
        var td = document.createElement( "td" );
        td.innerText = "Alias";
        tr.appendChild( td );
       
        var td = document.createElement( "td" );
        td.innerText = sheetinformation.alias;
        tr.appendChild( td );

        var td = document.createElement( "td" );
        td.innerText = "Description";
        tr.appendChild( td );
       
        var td = document.createElement( "td" );
        td.innerText = sheetinformation.description;
        tr.appendChild( td );

    }

    function createRecordsTable( columninformation, records ){

        var previewSheetsTable = document.getElementById( "previewSheetsTable" );


        var tr = document.createElement( "tr" );
        previewSheetsTable.appendChild( tr );

        var td = document.createElement( "td" );
        tr.appendChild( td );

        var table = document.createElement( "table" );
        td.appendChild( table );
        
        var tr = document.createElement( "tr" );
        table.appendChild( tr );

        for( var i = 0; i < columninformation.length; i++ ){

            var td = document.createElement( "td" );
            td.innerText = columninformation[i].nameInExcel;
            tr.appendChild( td );
           
        }

        for( var i = 0; i < records.length; i++ ){

            var tr = document.createElement( "tr" );
            table.appendChild( tr );

            var record = records[i];

            for( var j = 0; j < record.length; j++ ){


                var td = document.createElement( "td" );
                td.innerText = record[j];
                tr.appendChild( td );

            }

        }
        
    }

    function editSheet( inputId ){

        var sheetId = new Number( inputId.substring( "editSheet".length ) ).valueOf();

        for(var i = 0; i < sheetsObject.sheets.length; i++){

            var sheet = sheetsObject.sheets[i];

            if( sheet.id == sheetId ){

                editingSheet = sheet;

                break;
                
            }

        }

        document.getElementById( "editSheetAlias" ).value = editingSheet.alias;
        document.getElementById( "editSheetDescription" ).value = editingSheet.description;
        

    }
    
};

ExcelBrowser.prototype = {
};



