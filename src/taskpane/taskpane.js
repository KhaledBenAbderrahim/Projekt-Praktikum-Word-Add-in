/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
//! import "../../assets/logo.png";

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    
    /*document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("insert-table").onclick = insertTable;*/
   
    document.getElementById("button").onclick = button;


  }
});

function button() {
  if(document.getElementById("button").innerText=="Block Beginn"){

      blockBeginn()
      document.getElementById("button").innerText="Block Ende" 


  }
    else if(document.getElementById("button").innerText=="Block Ende"){

      blockEnde()
      document.getElementById("button").innerText="Block Beginn" 


  }

}

function insertParagraph() {
  document.getElementById("insert-paragraph").innerText="insert-name"
  Word.run(function (context) {

      // TODO1: Queue commands to insert a paragraph into the document.
      var docBody = context.document.body;
      docBody.insertParagraph("Zahlungsplan ", "Start");

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function insertTable() {
  Word.run(function (context) {

      // TODO1: Queue commands to get a reference to the paragraph
      //        that will proceed the table.
      var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
      // TODO2: Queue commands to create a table and populate it with data.
      var tableData = [
        ["Fällig am", "Fälliger Betrag", "Saldo"],
        ["122020", "1000", "0"],
        ["012021", "1000", "0"],
        ["022021", "1000", "-1000"],
      ];
      secondParagraph.insertTable(4, 3, "After", tableData);
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function blockBeginn() {
  Word.run(function (context) {

      // TODO1: Queue commands to insert a paragraph into the document.
      var docBody = context.document.body;
      docBody.insertText("${B:0} ", "End");

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function blockEnde() {
  Word.run(function (context) {

      // TODO1: Queue commands to insert a paragraph into the document.
      var docBody = context.document.body;
      docBody.insertText("${B:1} ", "End");

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}






