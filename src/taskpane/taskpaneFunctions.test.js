
const beginnBlock = require("./taskpaneFunctions")
import {BlockBeginn,blockEnde,newDisplay,button,addCondition} from './taskpaneFunctions'




describe("Taskpane.js Funktionen Testen", () => {
    let docBody , ctrl,insert,title ,tag, appearance,color,selection,Word,context,document,Office

    beforeAll(() => {
         window.Office = {
             onReady:() => {

             }
         }
         window.Word = {
            
            
            run: (resolve) => {
                resolve(context);
            
            },
            contentControl:{
                title ,
                tag ,
                appearance ,
                color,
                parentBody:{
                    
                        select:(selectionMode)=>{
                            selection = selectionMode
                            return selection
                        }
                
                },
            

            }
         }
        
        context = {
            sync: () => {new Promise(accept => accept());},
            document:{
                body: {
                    paragraphs: []
                  },
                load: () => '',
                getSelection: () => {
                    return {
                        
                        values: "25",
                        load: () => '',
                        insertHtml: (html,insertLocation) => {
                            return {
                                html,
                                insertLocation,
                                }
                            },
                        insertContentControl: ()=> {
                            return window.Word.contentControl
                        }
                        
                            };
                        },
                
            }
        }
        
        
        

})

    test("ob ein Block Beginnt",()=>{
        
        // #beginn.blockBeginn()
        BlockBeginn()
        expect(context.document.getSelection().values).toBe("25")
        expect(context.document.getSelection().insertHtml("${B:0}", "Start").html).toBe("${B:0}")
        expect(window.Word.contentControl.title).toBe("BlockBegin")
        //expect(Word.contentControl.title).toBe("BlockEnde")
        // expect(selection).toBe("End")
        })
    test("ob ein Block Endet",()=>{
        
            // #beginn.blockBeginn()
        blockEnde()
        expect(context.document.getSelection().values).toBe("25")
        expect(context.document.getSelection().insertHtml("${B:1}", "Start").html).toBe("${B:1}")
        expect(window.Word.contentControl.title).toBe("BlockEnde")
            //expect(Word.contentControl.title).toBe("BlockEnde")
            // expect(selection).toBe("End")
            })
    test("Button 'block hinzüfügen' testen",()=>{
        window.document.body.innerHTML = `
            <button type="button" class="btn btn-primary btn-sm form-control" id="button">Block Beginn</button>
            `;
        button()
        expect(window.document.getElementById("button").textContent).toBe("Block Ende")
        expect(window.document.getElementById("button").className).toBe("btn btn-danger btn-sm form-control")
    })
    test("die funktion newDisplay/'(von myFunction)' testen",()=>{
        var answer = newDisplay("none")
        expect(answer).toBe("block")
    })

    test("ob ein Alert gezeigt bei leeren Bedingung",()=>{
        window.document.body.innerHTML = `
        <div class="condition" id="myDIV" style="display:none" >
            
        <div class="input-group form-control">
            
            <input id="f1" placeholder="F1  auswaehlen" type="text" class="form-control" aria-label="Text input with radio button" >
          </div>
          <div class="form-control">
            <select class="form-control" id="condition" required>
              <option selected>Operator</option>
              <option value="1">=</option>
              <option value="2"><></option>
              <option value="3">></option>
              <option value="4"><</option>
              <option value="5">>=</option>
              <option value="6"><=</option>
              <div class="invalid-feedback">Example invalid custom select feedback</div>

            </select>
            
          </div>
          <div class="input-group mb-3 form-control">
            <div class="input-group-text">
                <input id="check" class="form-check-input mt-0" type="checkbox" value="" aria-label="Checkbox for following text input">
              </div>
                <input id="f2" placeholder="F2 auswaehlen(optional)" type="text" class="form-control" aria-label="Text input with checkbox">
          </div>
          <div class="form-control">
            <select class="form-control" id="action">
              <option selected>Aktion</option>
              <option value="0">0 = Kompletten Absatz löschen</option>
              <option value="1">1 = Verbleibende Platzhalter löschen </option>
              <option value="2">2 = Nächsten Platzhalter löschen</option>
              <option value="3">3 = Rest des Absatzes löschen</option>
              <option value="4">4 = Bis zum nächsten Platzhalter löschen</option>
            </select>
        
          </div>
          
          <div class="form-control">
          <button type="button" class="btn btn-success btn-sm form-control" id="create-condition">Bedingung hinzufügen</button>
          </div>
          <div id="alert" class="alert alert-danger" role="alert" style="display:none">
            Feld 1 , Bedingung und Aktion sind Pflichtfelder
          </div>
          
        

        `;
        addCondition();
        expect(window.document.getElementById("alert").style.display).toBe("block");
    })
})



