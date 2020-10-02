// exercise 2, create http service with a single post

// Since libraries are allowed I'm using express and body-parser. On a personal note, it's been a few months since I coded, 
// because of moving to a different country and the loong cuarentine in Argentina. I did have trouble with this excercise, ended up sending the json through Ajax
// but I didn't got the code to work, I have used ajax in the past and I am very familiar with http, please check my other repo: https://github.com/AlfredoThill/durin.monsters.

// setting up the modules and the app
const express = require("express"); 
const bodyParser = require("body-parser") 
const app = express(); 
app.use(bodyParser.urlencoded({ 
    extended:true
})); 


app.get("/", function(req, res) { 
  res.sendFile(__dirname + "/excercise2.html"); 
}); 

// The post
app.post("/palindrome", function(req, res) { 
  const phrase = req.body["phrase"]; 
  const boolean = palindrome(phrase);  
  res.send({
    "palindrome": boolean
  }); 
}); 
  
app.listen(3000, function(){ 
  console.log("server is running on port 3000"); 
}) 

// Auxiliar function with regular expresions to check for palindrome
function palindrome(str) {
    const re = /[\W_]/g;
    const lowRegStr = str.toLowerCase().replace(re, '');
    const reverseStr = lowRegStr.split('').reverse().join(''); 
    return reverseStr === lowRegStr;
}
