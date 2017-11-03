const express = require('express'),
  app = express(),
  auth = require('./App/Middleware/auth'),
  reload = require('reload'),
  cookieParser = require('cookie-parser'),
  cors = require('cors'),
  session = require('express-session'),
  helmet = require('helmet'),
  nunjucks = require('nunjucks'),
  microsoftGraph = require("@microsoft/microsoft-graph-client");

const c = console.log

nunjucks.configure('App/Views', {
  autoescape: true,
  express: app
});

function authorizeSession(req, res, next) { // Session authorization
  if (req.method === 'GET') {
    if (!req.session.o365AccessToken && req.path !== '/authorize') {
      res.redirect(auth.getAuthUrl());
      return;  
    } 
    if (new Date(parseFloat(req.session.o365TokenExpires)) <= new Date()) {
      let refreshToken = req.session.o365RefreshToken
      auth.refreshAccessToken(refreshToken, (err, newToken) => {
        if (err) console.log('Error: ' + err)
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
  cookie: {
    secure: false,
    expires: 1000 * 60 * 60 * 24
  } // Cookie expiration date set to 24H
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

function getMyEmail(token) {

  var client = microsoftGraph.Client.init({
    authProvider: (done) => {
      // Just return the token
      done(null, token);
    }
  });
  var result;
  client
    .api('/me/mailfolders/inbox/messages')
    .top(10)
    .select('subject,from,receivedDateTime,isRead')
    .orderby('receivedDateTime DESC')
    .get((error, response) => {
      if (error) {
        console.log(err);
      } else {
        return response
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

/////////ROUTES


app.get('/', function (req, res) {
  res.render('index.html', {
    title: "root page"
  })
})

app.get('/main', function (req, res) {
  res.render('index.html', {
    title: "main page"
  })
})

app.get('/authorize', (req, res) => {
  auth.getTokenFromCode(req.query.code, tokenReceived, req, res)
})

app.get('/logout', (req, res) => {
  req.session.destroy()
  res.render('index.html', {
    title: "logout"
  })
})

app.get('/mail', (req, res) => {
  c(getMyEmail(req.session.o365AccessToken))
  res.render('email.html', {
    title: "email",  
    emails: getMyEmail(req.session.o365AccessToken)
  })
})




reload(app)
app.listen(8080, console.log('Server running on localhost:8080'))