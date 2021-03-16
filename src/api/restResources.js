


const axios = require("axios");

function getAllData(){
    return axios.get("https://demo.akademie.uni-bremen.de/rest/meta")
                .then(function (response){
                    return ( response.data);
                })
}

getAllData().then(function (response){console.log(response);})

 