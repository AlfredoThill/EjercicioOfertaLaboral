// exercise 2, create http service with a single post using express

const express = require("express"); 
const bodyParser = require("body-parser") 
  
const app = express(); 
app.use(bodyParser.urlencoded({ 
    extended:true
})); 
  
app.get("/", function(req, res) { 
  res.sendFile(__dirname + "/excercise2.html"); 
//  res.send({ phrase: "mi frase"})
}); 
  
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

function palindrome(str) {
    const re = /[\W_]/g;
    const lowRegStr = str.toLowerCase().replace(re, '');
    const reverseStr = lowRegStr.split('').reverse().join(''); 
    return reverseStr === lowRegStr;
}
