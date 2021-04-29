



function BlockBeginn (){
    Word.run(function (context) {
  
        // TODO1: Queue commands to insert a paragraph into the document.
        let docBody = context.document.getSelection();
        docBody.insertHtml("${B:0}", "End");
        let ctrl = docBody.insertContentControl()
        ctrl.title = "BlockBegin";
        ctrl.tag = "Select";
        ctrl.appearance = "BoundingBox";
        ctrl.color = "#589CFB";
        ctrl.parentBody.select("End")
  
        return context.sync();
    })
}



export{BlockBeginn,blockEnde,newDisplay,button,addCondition}

function blockEnde() {
    Word.run(function (context) {
  
        // TODO1: Queue commands to insert a paragraph into the document.
        let docBody = context.document.getSelection();
        docBody.insertHtml("${B:1} ", "End");
        let ctrl = docBody.insertContentControl()
        ctrl.title = "BlockEnde";
        ctrl.tag = "Select";
        ctrl.appearance = "BoundingBox";
        ctrl.color = "#589CFB";
        ctrl.parentBody.select("End")
  
  
        return context.sync();
    })
}

function newDisplay(display){
    if (display === "none") {
      display = "block";
    } else {
      display = "none";
    }
    return display;
  }

  function button() {
    if(window.document.getElementById("button").textContent=="Block Beginn"){
  
        BlockBeginn()
        window.document.getElementById("button").textContent="Block Ende"
        window.document.getElementById("button").className="btn btn-danger btn-sm form-control"
        
  
  
    }
      else if(document.getElementById("button").textContent=="Block Ende"){
  
        blockEnde()
        window.document.getElementById("button").textContent="Block Beginn"
        window.document.getElementById("button").className="btn btn-primary btn-sm form-control"
        }
  
  }

  function addCondition(){
    var action = window.document.getElementById("action");
    var condition = window.document.getElementById("condition");

    var checked = window.document.getElementById("check").checked;
    var feld1Input = window.document.getElementById("f1").value;
    var feld2Input = checked?window.document.getElementById("f2").value:"" ;

    var actionValue = action.options[action.selectedIndex].value;
    var actionResult =  ":" + actionValue;
    var conditionResult =  condition.options[condition.selectedIndex].text;
    var x = window.document.getElementById("alert")
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
  });
      x.style.display = "none"
}else{
      x.style.display = "block"
}

}

  




    

