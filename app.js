const express = require('express'),
  app = express(),
  auth = require('./App/Middleware/auth'),
  reload = require('reload'),
  cookieParser = require('cookie-parser'),
  cors = require('cors'),
  microsoftGraph = require("@microsoft/microsoft-graph-client");

app.use(cookieParser())
app.use(cors())
  
const c = console.log

function getUserEmail(token, callback) {
  // Create a Graph client
  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      // Just return the token
      done(null, token);
    }
  });

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

function getMyEmail(token, res) {

  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      // Just return the token
      done(null, token);
    }
  });

  client
  .api('/me/mailfolders/inbox/messages')
  .top(10)
  .select('subject,from,receivedDateTime,isRead')
  .orderby('receivedDateTime DESC')
  .get((error, response) => {
    if (error) {
      console.log(err);
    } else {
      console.log(response)
    }
  })
}

function tokenReceived(res, error, token) {
  if (error) {
    console.log('Access token error: ', error.message);
  } else {
    res.cookie('o365AccessToken', token.token.access_token, {expire: new Date() + 4000})
    res.cookie('o365RefreshToken', token.token.refresh_token, {expire: new Date() + 4000})
    res.cookie('o365TokenExpires', token.token.expires_at.getTime()).send('initial cookies')
    
  }
}

app.get('/', function (req, res) {
  res.redirect(auth.getAuthUrl());
})

app.get('/main', function (req, res) {
  c(req.cookies)
  res.send('<a href="/mail">renew cookies</a>')
})


app.get('/authorize', (req, res) => {
  auth.getTokenFromCode(req.query.code, tokenReceived, res)
})

app.get('/mail', (req, res) => {
  c(req.cookies)
  if (true ) {
    c('token expired')
    let refreshToken = req.cookies.o365RefreshToken
    auth.refreshAccessToken(refreshToken, (err, newToken) => {
      if (err) console.log('Error: '+err) 
      else if (newToken) {
        res.cookie('o365AccessToken', newToken.token.access_token, {expire: new Date() + 4000})
        res.cookie('o365RefreshToken', newToken.token.refresh_token, {expire: new Date() + 4000})
        res.cookie('o365TokenExpires', newToken.token.expires_at.getTime(), {expire: new Date() + 4000}).send('new cookies')        
      }
    })
  }
})




reload(app)
app.listen(8080, console.log('Server running on localhost:8080'))