# Introduction

This mock will help you to write the unit test cases for the **office.js** application using **karma**. While testing office.js application its has Office and Excel as global object. this package help to get those variable declared for you.


## Installaion

    npm i officejs-mock --save
    
## Referencing
update  karma.config.ts file to load the mock file

    files: ['../node_modules/officejs-mock/mock.ts'],
also add mime type as follows

    mime: {'text/x-typescript': ['ts']},

## Usage

In this package all the function available over office.js is define to support the unit testing as follows 
```javascript
    window['context'] = {
	    workbook: {
		    getSelectedRange : () => {
			    return {
				    load: () =>  '',
				    values:  window['values'] || [],
				    cellCount:  window['cellCount'] ||  0
			    };
		    }
	    }
    };
```
in which the referencing data has been returned from the window object which help to configure the data while writing test cases 

Ex.
consider we are writing test case for the function which return the selected cell value 
```javascript
    async getSelectedCellValue(){
	    let cellValue;
        await Excel.run(async context => {
	          const range = context.workbook.getSelectedRange();
	          range.load();
	          await context.sync();
	          cellValue = range.values;
        });
        return cellValue;
    }
```

spec.ts will be look like 

```javascript
    it('getSelectedCellValue function should get the selected cell text', async(async () => {
            // tslint:disable-next-line: no-string-literal
            window['values'] = [['cellValue']];
            expect(await app.getSelectedCellValue()).toEqual([['cellValue']]);
    }));
```

window['values']  is used to configure your test date which will be used by mock to return it.


## List of Global variable :

|  Variable  |  use|
|--|--|
| window['values'] |  Is for the selected rage values |
| window['cellCount'] |  Is for the selected rage cellCount |
| window['chartId'] | Is to get the selected chartid on onActivated even |



