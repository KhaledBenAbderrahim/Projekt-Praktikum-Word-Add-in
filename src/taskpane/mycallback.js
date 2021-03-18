"use strict"
const axios = require("axios");

function getAllData(){
    return axios.get("https://demo.akademie.uni-bremen.de/rest/meta?jsonp=acat")
                .then(function (response){
                      return response.data;
                      

                     

                })
}


getAllData().then(function (response){console.log(response)})