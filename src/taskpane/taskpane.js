
 
// images references in the manifest
// import "../../assets/icon-16.png";
// import "../../assets/icon-32.png";
// import "../../assets/icon-80.png";
//! import "../../assets/logo.png";

/* global document, Office, Word */

/**
 * Office.onReady()ist eine asynchrone Methode. 
 * @return {Promise-Objekt}  
 * Ein Promise-Objekt zurückgibt, während überprüft wird, ob die Office.js-Bibliothek geladen ist.
 * Wenn die Bibliothek Office.js geladen wird, löst es das Versprechen als ein Objekt, das die Office-Client-Anwendung mit  
 * einem angibt Office.HostTypeENUM - Wert( Excel, Word, etc.) und die Plattform mit einem Office.PlatformTypeENUM - Wert( PC, Mac, OfficeOnline, etc.). 
 * Das Versprechen wird sofort aufgelöst, wenn die Bibliothek beim Office.onReady()Aufruf bereits geladen ist.
 * durch diese funktion , können wir unsere taskpane Bedienen durch die drei Methoden:button(),myFunction(),
 * addCondition()
 */

window.Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    
    /*document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("insert-table").onclick = insertTable;*/
   
    document.getElementById("button").onclick = button;
    document.getElementById("hide/show").onclick = myFunction;
    document.getElementById("create-condition").onclick = addCondition;
    


  }
});



/**
 * Die Funktion erstellt ein Block und zeigt "Block Beginn" oder "Block Ende".
 * @return {element} button mit text "Block Beginn"/"Block Ende" 
 */
function button() {
  if(document.getElementById("button").innerText=="Block Beginn"){

      blockBeginn()
      document.getElementById("button").innerText="Block Ende"
      document.getElementById("button").className="btn btn-danger btn-sm form-control"
      


  }
    else if(document.getElementById("button").innerText=="Block Ende"){

      blockEnde()
      document.getElementById("button").innerText="Block Beginn"
      document.getElementById("button").className="btn btn-primary btn-sm form-control"
      }

}




/**
 * Die Funktion schreibt ein String '${B:0}'(BlockBeginn) und erzeugt ein ContentControl auf diese String Element in MS Word..
 * 
 * Mit hilfe der Funktion Word.run() führt ein Stapelskript aus, das mithilfe des neuen RequestContext bestimmte Aktionen 
 * für die Word-Objektmodelle ausführt. Seit dem Office-Add-In und dem Word Die Anwendung wird in zwei verschiedenen Prozessen 
 * ausgeführt. Die im Anforderungskontext verfügbare sync()-Methode synchronisiert den Status zwischen den JavaScript-Proxyobjekten 
 * und den realen Objekten in der Office-Anwendung.
 * @param {RequestContext} - Wir müssen RequestContext übergeben, um auf Word-Objektmodelle zugreifen zu können.
 * @return {context.sync()} - Sendet die Anforderungswarteschlange an die Office-Anwendung und gibt ein Versprechungsobjekt zurück,
 * mit dem weitere Aktionen.
 * @throws  {error} - Wenn ein Versprechen aufgrund eines Fehlers bei der Verarbeitung der Anforderung abgelehnt wird.
 */
function blockBeginn() {
  Word.run(function (context) {

      // TODO1: Queue commands to insert a paragraph into the document.
      let docBody = context.document.getSelection();
      docBody.insertHtml("${B:0} ", "End");
      let ctrl = docBody.insertContentControl()
      ctrl.title = "BlockBegin";
      ctrl.tag = "Select";
      ctrl.appearance = "BoundingBox";
      ctrl.color = "#589CFB";
      ctrl.parentBody.select("End")

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

/**
 * Die Funktion schreibt ein String '${B:1}'(BlockEnde) und erzeugt ein ContentControl auf diese String Element in MS Word..
 * 
 * Mit hilfe der Funktion Word.run() führt ein Stapelskript aus, das mithilfe des neuen RequestContext bestimmte Aktionen 
 * für die Word-Objektmodelle ausführt. Seit dem Office-Add-In und dem Word Die Anwendung wird in zwei verschiedenen Prozessen 
 * ausgeführt. Die im Anforderungskontext verfügbare sync()-Methode synchronisiert den Status zwischen den JavaScript-Proxyobjekten 
 * und den realen Objekten in der Office-Anwendung.
 * @param {RequestContext} - Wir müssen RequestContext übergeben, um auf Word-Objektmodelle zugreifen zu können.
 * @return {context.sync()} - Sendet die Anforderungswarteschlange an die Office-Anwendung und gibt ein Versprechungsobjekt zurück,
 * mit dem weitere Aktionen.
 * @throws  {error} - Wenn ein Versprechen aufgrund eines Fehlers bei der Verarbeitung der Anforderung abgelehnt wird.
 */

function blockEnde() {
  Word.run(function (context) {

      // TODO1: Queue commands to insert a paragraph into the document.
      var docBody = context.document.getSelection();
      docBody.insertHtml("${B:1} ", "End");
      const ctrl = docBody.insertContentControl()
      ctrl.title = "BlockEnde";
      ctrl.tag = "Select";
      ctrl.appearance = "BoundingBox";
      ctrl.color = "#589CFB";
      ctrl.parentBody.select("End")


      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

/**
 * Der wert von ein String wechseln je nach Kondition (none/block)
 * @param {String} display 
 * @return {String} 
 */
function newDisplay(display){
  if (display === "none") {
    display = "block";
  } else {
    display = "none";
  }
  return display;
}

/**
 * mit hilfe diese Funktion, Ein Bedingung Formular wird gezeigt und verbergt mit hilfe der zürückgegebene Wert von
 * newDisplay() methode.
 * @return {element} none - nicht einzeigen
 * @return {element} block - einzeigen
 */
function myFunction() {
  var x = document.getElementById("myDIV");
  console.log("Button gedrückt");
  console.log("Wert: " +x.style.display);
  var display = x.style.display;
  display = newDisplay(display);
  x.style.display = display;
}


/**
 * Die Funktion schreibt ein String (bedingung) und erzeugt ein ContentControl auf diese String Element in MS Word..
 * Diese String wird gebaut durch die inhaltene werte von ein Formular 
 * Mit hilfe der Funktion Word.run() führt ein Stapelskript aus, das mithilfe des neuen RequestContext bestimmte Aktionen 
 * für die Word-Objektmodelle ausführt. Seit dem Office-Add-In und dem Word Die Anwendung wird in zwei verschiedenen Prozessen 
 * ausgeführt. Die im Anforderungskontext verfügbare sync()-Methode synchronisiert den Status zwischen den JavaScript-Proxyobjekten 
 * und den realen Objekten in der Office-Anwendung.
 * @param {RequestContext} - Wir müssen RequestContext übergeben, um auf Word-Objektmodelle zugreifen zu können.
 * @return {context.sync()} - Sendet die Anforderungswarteschlange an die Office-Anwendung und gibt ein Versprechungsobjekt zurück,
 * mit dem weitere Aktionen.
 * @throws  {error} - Wenn ein Versprechen aufgrund eines Fehlers bei der Verarbeitung der Anforderung abgelehnt wird.
 */
function addCondition(){
      var action = document.getElementById("action");
      var condition = document.getElementById("condition");

      var checked = document.getElementById("check").checked;
      var feld1Input = document.getElementById("f1").value;
      var feld2Input = checked?document.getElementById("f2").value:"" ;

      var actionValue = action.options[action.selectedIndex].value;
      var actionResult =  ":" + actionValue;
      var conditionResult =  condition.options[condition.selectedIndex].text;
      var x = document.getElementById("alert")
      if(feld1Input.length>0 && actionValue!="Aktion" && conditionResult!="Operator" ){
      Word.run(function (context) {

        // TODO1: Queue commands to insert a paragraph into the document.
        var docBody = context.document.getSelection();
        docBody.insertHtml("${C:"+feld1Input+":"+conditionResult+feld2Input+actionResult+"}" , "Start");
        const ctrl = docBody.insertContentControl()
        ctrl.title = "Bedingung";
        ctrl.tag = "Select";
        ctrl.appearance = "BoundingBox";
        ctrl.color = "#589CFB";
        ctrl.parentBody.select("End")


        return context.sync();
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
    x.style.display = "none"
  }else{
    x.style.display = "block"
  }

}

function onChangedTree(e, data) {
  var x = document.getElementById("myDIV");
  if (x.style.display === "none") {
    Word.run(function(context) {
      var docBody = context.document.getSelection();
      docBody.insertHtml("${F:" + data.selected + "}", Word.InsertLocation.end);
      const ctrl = docBody.insertContentControl();
      ctrl.title = "Select";
      ctrl.tag = "Select";
      ctrl.appearance = "BoundingBox";
      ctrl.color = "#589CFB";
      ctrl.parentBody.select("End");

      return context.sync();
    });
  } else {
    if (!$("#check").is(":checked")) {
      document.getElementById("f1").value = data.selected;
    }
    if ($("#check").is(":checked")) {
      document.getElementById("f2").value = "F" + data.selected;
    }
  }
}






