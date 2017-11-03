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

// Session authorization

function authorizeSession (req, res, next) {
  if (req.method === 'GET') { 
    if (!req.session.o365AccessToken && req.path!=='/authorize') res.redirect(auth.getAuthUrl());
  }
  // keep executing the router middleware
  next()
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
  cookie: { secure: false, expires: 360000000 }
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
    c(req)
    res.redirect('/main')
    // res.cookie('o365AccessToken', token.token.access_token, {maxAge: cookieExpirationTime})
    // res.cookie('o365RefreshToken', token.token.refresh_token, {maxAge: cookieExpirationTime})
    // res.cookie('o365TokenExpires', token.token.expires_at.getTime(), {maxAge: cookieExpirationTime}).send('initial cookies')
    
  }
}
app.get('/', function (req, res) {
  // if (req.session.o365AccessToken) res.send('got sessions') 
  // else res.redirect(auth.getAuthUrl());
  // res.send('/')
})

app.get('/main', function (req, res) {
  c(req.session)
  res.send('<a href="/mail">renew cookies</a>')
})


app.get('/authorize', (req, res) => {
  auth.getTokenFromCode(req.query.code, tokenReceived, req, res)
})

app.get('/logout', (req, res) => {
  req.session.destroy()
})

app.get('/mail', (req, res) => {
  c(req.cookies)
  if (true ) {
    c('token expired')
    let refreshToken = req.cookies.o365RefreshToken
    auth.refreshAccessToken(refreshToken, (err, newToken) => {
      if (err) console.log('Error: '+err) 
      else if (newToken) {
        res.cookie('o365AccessToken', newToken.token.access_token, {maxAge: cookieExpirationTime})
        res.cookie('o365RefreshToken', newToken.token.refresh_token, {maxAge: cookieExpirationTime})
        res.cookie('o365TokenExpires', newToken.token.expires_at.getTime(), {maxAge: cookieExpirationTime}).send('new cookies')        
      }
    })
  }
})




reload(app)
app.listen(8080, console.log('Server running on localhost:8080'))