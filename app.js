const express = require('express'),
  app = express(),
  auth = require('./auth'),
  reload = require('reload')
  microsoftGraph = require("@microsoft/microsoft-graph-client");
  
const reloadTag =  '<script src="/reload/reload.js"></script>';  

const c = console.log

////////////////////////////////////////////

function htmlContent(res) {
  let htmlContentType = ['Content-Type', 'text/html']
  res.header.apply(htmlContentType)
}

/////////////////////////////////////////

function getUserEmail(token, callback) {
  // Create a Graph client
  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      // Just return the token
      done(null, token);
    }
  });

  // Get the Graph /Me endpoint to get user email address
  client
    .api('/me')
    .get((err, res) => {
      if (err) {
        callback(err, null);
      } else {
        callback(null, res.mail);
      }
    });
}

function tokenReceived(response, error, token) {
  if (error) {
    console.log('Access token error: ', error.message);
  } else {
    getUserEmail(token.token.access_token, function(error, email) {
      if (error) {
        console.log('getUserEmail returned an error: ' + error);
      } else if (email) {
        c(email)
      }
    });
  }
}

/////////////////////////////////////////


app.get('/', function (req, res) {
  htmlContent(res)
  res.send(`<a href='${auth.getAuthUrl()}'>Sign in Here</a>`)
})

app.get('/authorize', (req, res) => {
  auth.getTokenFromCode(req.query.code, tokenReceived, res)
})


//////////////////////////////////////////

reload(app)

const port = 8080
app.listen(port, () => console.log('Server running on localhost:' + port))