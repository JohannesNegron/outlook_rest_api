'use strict'
const http = require('http');
const app = require('express')();
var request = require("request");
const port = 3000;

let client_id = '3be0d09a-9b47-44df-94f8-4112171d215a';
let client_secret = 'usqfinqLL3532)^(MYORI4)';
let redirect1 = 'http://localhost:3000/auth'

app.get('/', (req, res)=>
{
    res.redirect('https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id='+client_id+'&response_type=code&redirect_uri=http://localhost:3000/callback&scope=User.Read')
})

app.get('/callback', (req, res)=>
{
    console.log('llego a callback')
    const code = req.query.code;

    var options = { 
        method: 'POST',
        url: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        form: 
        { 
            grant_type: 'authorization_code',
            code: code,
            redirect_uri: 'http://localhost:3000/callback',
            client_id: client_id,
            client_secret: client_secret 
        },
        headers: {
            'content-type': 'application/x-www-form-urlencoded' 
        } 
    };

    request(options, function (error, response, body) 
    {
        if (error) throw new Error(error);
        console.log(body);
        res.send(body);
    });

    
})
app.get('/auth',(req, res)=>
{
    console.log(req.query)
})

app.listen(3000, (err)=>
{
    if(err) return console.error(err);
    console.log(`Express server listening at http://localhost:${port}`)
});
