var credentials = {
    client: {
        id: '9919e965-ffb2-4b44-97fc-e855c7f28c69',
        secret: 'dxavkHFGQ491_!^amLQF88~',
    },
    auth: {
        tokenHost: 'https://login.microsoftonline.com',
        authorizePath: 'common/oauth2/v2.0/authorize',
        tokenPath: 'common/oauth2/v2.0/token'
    }
};

/////CONFIG

var oauth2 = require('simple-oauth2').create(credentials);

var redirectUri = 'http://localhost:8080/authorize';

// The scopes the app requires
var scopes = ['openid', 'User.Read', 'Mail.Read'];

function getAuthUrl() {
    var returnVal = oauth2.authorizationCode.authorizeURL({
        redirect_uri: redirectUri,
        scope: scopes.join(' ')
    });
    console.log('Generated auth url: ' + returnVal);
    return returnVal;
}


function getTokenFromCode(auth_code, callback, response) {
    var token;
    oauth2.authorizationCode.getToken({
        code: auth_code,
        redirect_uri: redirectUri,
        scope: scopes.join(' ')
    }, function (error, result) {
        if (error) {
            console.log('Access token error: ', error.message);
            callback(response, error, null);
        } else {
            token = oauth2.accessToken.create(result);
            console.log('Token created: ', token.token);
            callback(response, null, token);
        }
    });
}

exports.getTokenFromCode = getTokenFromCode;
exports.getAuthUrl = getAuthUrl;