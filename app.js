const express = require('express')
const app = express()
const auth = require('./auth')
const microsoftGraph = require("@microsoft/microsoft-graph-client");

const c = console.log

function htmlContent (res) {
  let htmlContentType = ['Content-Type', 'text/html']
  res.header.apply(htmlContentType)
}

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


app.get('/', function (req, res) {
  htmlContent(res)
  res.send(`<a href='${auth.getAuthUrl()}'>Sign in Here</a>`)
})

app.get('/authorize', (req, res) => {
  htmlContent(res)
  res.send(`<h1>authorize page - code is: ${req.query.code}</h1>`)
  console.log(req.query.code)
})



const port = 8080
app.listen(port, () => console.log('Server running on localhost:' + port))