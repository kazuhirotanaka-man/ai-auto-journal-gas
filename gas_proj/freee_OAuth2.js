/**
 * @OnlyCurrentDoc
 *
 * The MIT License (MIT)
 *
 * Copyright (c) 2015-2018 Googledrive
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *
 * @license MIT
 */
var OAuth2 = (function() {
  'use strict';

  // oob (Out-of-Band) a.k.a. Manual Copy/Paste
  var OOB_URI = 'urn:ietf:wg:oauth:2.0:oob';

  function createService(serviceName) {
    return new Service(serviceName);
  }

  function Service(serviceName) {
    this.serviceName_ = serviceName;
    this.params_ = {};
    this.scope_ = [];
  }

  Service.prototype.setAuthorizationBaseUrl = function(url) {
    this.authorizationBaseUrl_ = url;
    return this;
  };

  Service.prototype.setTokenUrl = function(url) {
    this.tokenUrl_ = url;
    return this;
  };

  Service.prototype.setClientId = function(id) {
    this.clientId_ = id;
    return this;
  };

  Service.prototype.setClientSecret = function(secret) {
    this.clientSecret_ = secret;
    return this;
  };

  Service.prototype.setPropertyStore = function(store) {
    this.propertyStore_ = store;
    return this;
  };
  
  Service.prototype.setScope = function(scope) {
    this.scope_ = Array.isArray(scope) ? scope : scope.split(' ');
    return this;
  };

  Service.prototype.getAuthorizationUrl = function() {
    var params = {
      'client_id': this.clientId_,
      'response_type': 'code',
      'redirect_uri': OOB_URI,
      'scope': this.scope_.join(' ')
    };
    for (var param in this.params_) {
      params[param] = this.params_[param];
    }
    var paramString = Object.keys(params).map(function(key) {
      return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
    }).join('&');
    return this.authorizationBaseUrl_ + '?' + paramString;
  };

  Service.prototype.exchangeCodeForToken = function(code) {
    var payload = {
      'code': code,
      'client_id': this.clientId_,
      'client_secret': this.clientSecret_,
      'redirect_uri': OOB_URI,
      'grant_type': 'authorization_code'
    };
    var options = {
      'method': 'post',
      'payload': payload,
      'muteHttpExceptions': true
    };
    var response = UrlFetchApp.fetch(this.tokenUrl_, options);
    var result = JSON.parse(response.getContentText());
    if (response.getResponseCode() !== 200) {
      throw new Error('Token exchange failed: ' + response.getContentText());
    }
    this.saveToken_(result);
    return true;
  };

  Service.prototype.hasAccess = function() {
    var token = this.getToken_();
    if (!token) {
      return false;
    }
    if (token.expires_in && this.isExpired_(token)) {
      try {
        this.refresh();
      } catch (e) {
        this.reset();
        return false;
      }
    }
    return true;
  };

  Service.prototype.getAccessToken = function() {
    if (!this.hasAccess()) {
      throw new Error('Access not granted. Please authorize.');
    }
    return this.getToken_().access_token;
  };

  Service.prototype.reset = function() {
    this.propertyStore_.deleteProperty(this.getPropertyKey_());
  };

  Service.prototype.refresh = function() {
    var token = this.getToken_();
    if (!token.refresh_token) {
      throw new Error('No refresh token found.');
    }
    var payload = {
      'refresh_token': token.refresh_token,
      'client_id': this.clientId_,
      'client_secret': this.clientSecret_,
      'grant_type': 'refresh_token'
    };
    var options = {
      'method': 'post',
      'payload': payload,
      'muteHttpExceptions': true
    };
    var response = UrlFetchApp.fetch(this.tokenUrl_, options);
    var newToken = JSON.parse(response.getContentText());
    if (response.getResponseCode() !== 200) {
      throw new Error('Refresh failed: ' + response.getContentText());
    }
    this.saveToken_(newToken);
  };

  Service.prototype.saveToken_ = function(token) {
    token.granted_time = new Date().getTime();
    this.propertyStore_.setProperty(this.getPropertyKey_(), JSON.stringify(token));
  };

  Service.prototype.getToken_ = function() {
    var tokenString = this.propertyStore_.getProperty(this.getPropertyKey_());
    return tokenString ? JSON.parse(tokenString) : null;
  };

  Service.prototype.isExpired_ = function(token) {
    var now = new Date().getTime();
    var elapsed = (now - token.granted_time) / 1000;
    return elapsed >= token.expires_in;
  };

  Service.prototype.getPropertyKey_ = function() {
    return 'oauth2.' + this.serviceName_;
  };

  return {
    createService: createService
  };
})();
