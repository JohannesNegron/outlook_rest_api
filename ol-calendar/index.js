var sprLib = require('sprestlib');
var fs = require('fs');
var https = require('https'); // this Library is the basis for the remote auth solution
var CONTS = require('./utils');
var outlook = require("node-outlook");

class outlook_calendar
{
    constructor(user, pass, url_db, site)
    {
        this.sprLib = sprLib;
        this.sprLib.nodeConfig({nodeEnable:true});
        this.SP_USER = user;
        this.SP_PASS = pass;
        this.SP_URL = url_db;
        this.SP_HOST = url_db.toLowerCase().replace('https://','').replace('http://','');
        this.gBinarySecurityToken = "";
        this.gAuthCookie1 = "";
        this.gAuthCookie2 = "";
        this.xmlRequest = CONTS.xmlRequestFunction(user, pass, url_db);
        this.requestOptions = CONTS.requestOptions(this.xmlRequest.length)
        this.site = site;
    }
    print()
    {
        console.log(this)
    }
    auth()
    {
        let self = this;
        return new Promise((resolve,reject)=>
        {
            console.log(' * STEP 1/2: Auth into login.microsoftonline.com ...');
            return new Promise((resolve, reject)=>
            {
                var request = https.request(self.requestOptions, (res) => 
                {
                    let rawData = '';
                    res.setEncoding('utf8');
                    res.on('data', (chunk) => rawData += chunk);
                    res.on('end', () => {
                        var DOMParser = require('xmldom').DOMParser;
                        var doc = new DOMParser().parseFromString(rawData, "text/xml");
                        // KEY 1: Get SecurityToken
                        if ( doc.documentElement.getElementsByTagName('wsse:BinarySecurityToken').item(0) ) 
                        {
                            self.gBinarySecurityToken = doc.documentElement.getElementsByTagName('wsse:BinarySecurityToken').item(0).firstChild.nodeValue;
                            resolve();
                        }
                        else 
                        {
                            reject('Invalid Username/Password');
                        }
                    });
                });
    
                request.on('error', (err) =>
                {
                    console.log(`problem with request: ${err.message}`);
                    reject();
                });
                request.write(self.xmlRequest);
                request.end();
            })
            .then(()=>
            {
                var queryParams = {
                    '$select': 'Subject,ReceivedDateTime,From',
                    '$orderby': 'ReceivedDateTime desc',
                    '$top': 10
                  };
                outlook.mail.getMessage({token:self.gBinarySecurityToken, odataParams: queryParams},function(error, result)
                {
                    if (error) {
                        console.log('getMessages returned an error: ' + error);
                      }
                      else if (result) {
                        console.log('getMessages returned ' + result.value.length + ' messages.');
                        result.value.forEach(function(message) {
                          console.log('  Subject: ' + message.Subject);
                          var from = message.From ? message.From.EmailAddress.Name : "NONE";
                          console.log('  From: ' + from);
                          console.log('  Received: ' + message.ReceivedDateTime.toString());
                        });
                      }
                });
            })
        })
        
        .catch((strErr)=>
        {
            console.error('E R R O R');
            console.error(strErr)
            return;
        })
    }   
    
    userinfo()
    {
        return this.auth().then((data)=>
        {
            return data.user().info();
        });
    }
    
    
    consult(database_name, condition="")
    {
        return this.auth().then((data)=>
        {
            //console.log(data.list(database_name).getItems()).filter("ID eq 132")
            return data.list(database_name).items({queryFilter:condition})
        })
    }
    
}

module.exports = outlook_calendar;