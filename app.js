const express = require('express'),
  app = express(),
  auth = require('./App/Middleware/auth'),
  reload = require('reload'),
  cookieParser = require('cookie-parser'),
  cors = require('cors'),
  session = require('express-session'),
  helmet = require('helmet'),
  microsoftGraph = require("@microsoft/microsoft-graph-client");

const c = console.log


function authorizeSession (req, res, next) { // Session authorization
  if (req.method === 'GET') { 
    if (!req.session.o365AccessToken && req.path!=='/authorize') res.redirect(auth.getAuthUrl());
    if (new Date(parseFloat(req.session.o365TokenExpires)) <= new Date()) {
      let refreshToken = req.session.o365RefreshToken
      auth.refreshAccessToken(refreshToken, (err, newToken) => {
        if (err) console.log('Error: '+err) 
        else if (newToken) {
          req.session.o365AccessToken = newToken.token.access_token
          req.session.o365RefreshToken = newToken.token.refresh_token
          req.session.o365TokenExpires = newToken.token.expires_at.getTime()
        }
      })
    }
  }
  next() // keep executing the router middleware
}

app.use(cookieParser())
app.use(helmet())
app.disable('x-powered-by')
app.use(cors())
app.use(session({
  secret: 'smajdova manka',
  name: 'sessionID',
  httpOnly: false,
  resave: false,
  saveUninitialized: false,
  cookie: { secure: false, expires: 1000 * 60 * 60 * 24 } // Cookie expiration date set to 24H
}))
app.use(authorizeSession)

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

function tokenReceived(req, error, token, res) {
  if (error) {
    console.log('Access token error: ', error.message);
  } else {
    req.session.o365AccessToken = token.token.access_token
    req.session.o365RefreshToken = token.token.refresh_token
    req.session.o365TokenExpires = token.token.expires_at.getTime()
    res.redirect('/main')
  }
}
app.get('/', function (req, res) {

})

app.get('/main', function (req, res) {
  c(new Date(parseFloat(req.session.o365TokenExpires)))
  res.send('<a href="/mail">renew cookies</a>')
})

app.get('/authorize', (req, res) => {
  auth.getTokenFromCode(req.query.code, tokenReceived, req, res)
})

app.get('/logout', (req, res) => {
  req.session.destroy()
})

app.get('/mail', (req, res) => {

})




reload(app)
app.listen(8080, console.log('Server running on localhost:8080'))