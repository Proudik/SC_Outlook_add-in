/******/ (function() { // webpackBootstrap
/******/ 	var __webpack_modules__ = ({

/***/ "./node_modules/es6-promise/dist/es6-promise.js":
/*!******************************************************!*\
  !*** ./node_modules/es6-promise/dist/es6-promise.js ***!
  \******************************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

/*!
 * @overview es6-promise - a tiny implementation of Promises/A+.
 * @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
 * @license   Licensed under MIT license
 *            See https://raw.githubusercontent.com/stefanpenner/es6-promise/master/LICENSE
 * @version   v4.2.8+1e68dce6
 */

(function (global, factory) {
	 true ? module.exports = factory() :
	0;
}(this, (function () { 'use strict';

function objectOrFunction(x) {
  var type = typeof x;
  return x !== null && (type === 'object' || type === 'function');
}

function isFunction(x) {
  return typeof x === 'function';
}



var _isArray = void 0;
if (Array.isArray) {
  _isArray = Array.isArray;
} else {
  _isArray = function (x) {
    return Object.prototype.toString.call(x) === '[object Array]';
  };
}

var isArray = _isArray;

var len = 0;
var vertxNext = void 0;
var customSchedulerFn = void 0;

var asap = function asap(callback, arg) {
  queue[len] = callback;
  queue[len + 1] = arg;
  len += 2;
  if (len === 2) {
    // If len is 2, that means that we need to schedule an async flush.
    // If additional callbacks are queued before the queue is flushed, they
    // will be processed by this flush that we are scheduling.
    if (customSchedulerFn) {
      customSchedulerFn(flush);
    } else {
      scheduleFlush();
    }
  }
};

function setScheduler(scheduleFn) {
  customSchedulerFn = scheduleFn;
}

function setAsap(asapFn) {
  asap = asapFn;
}

var browserWindow = typeof window !== 'undefined' ? window : undefined;
var browserGlobal = browserWindow || {};
var BrowserMutationObserver = browserGlobal.MutationObserver || browserGlobal.WebKitMutationObserver;
var isNode = typeof self === 'undefined' && typeof process !== 'undefined' && {}.toString.call(process) === '[object process]';

// test for web worker but not in IE10
var isWorker = typeof Uint8ClampedArray !== 'undefined' && typeof importScripts !== 'undefined' && typeof MessageChannel !== 'undefined';

// node
function useNextTick() {
  // node version 0.10.x displays a deprecation warning when nextTick is used recursively
  // see https://github.com/cujojs/when/issues/410 for details
  return function () {
    return process.nextTick(flush);
  };
}

// vertx
function useVertxTimer() {
  if (typeof vertxNext !== 'undefined') {
    return function () {
      vertxNext(flush);
    };
  }

  return useSetTimeout();
}

function useMutationObserver() {
  var iterations = 0;
  var observer = new BrowserMutationObserver(flush);
  var node = document.createTextNode('');
  observer.observe(node, { characterData: true });

  return function () {
    node.data = iterations = ++iterations % 2;
  };
}

// web worker
function useMessageChannel() {
  var channel = new MessageChannel();
  channel.port1.onmessage = flush;
  return function () {
    return channel.port2.postMessage(0);
  };
}

function useSetTimeout() {
  // Store setTimeout reference so es6-promise will be unaffected by
  // other code modifying setTimeout (like sinon.useFakeTimers())
  var globalSetTimeout = setTimeout;
  return function () {
    return globalSetTimeout(flush, 1);
  };
}

var queue = new Array(1000);
function flush() {
  for (var i = 0; i < len; i += 2) {
    var callback = queue[i];
    var arg = queue[i + 1];

    callback(arg);

    queue[i] = undefined;
    queue[i + 1] = undefined;
  }

  len = 0;
}

function attemptVertx() {
  try {
    var vertx = Function('return this')().require('vertx');
    vertxNext = vertx.runOnLoop || vertx.runOnContext;
    return useVertxTimer();
  } catch (e) {
    return useSetTimeout();
  }
}

var scheduleFlush = void 0;
// Decide what async method to use to triggering processing of queued callbacks:
if (isNode) {
  scheduleFlush = useNextTick();
} else if (BrowserMutationObserver) {
  scheduleFlush = useMutationObserver();
} else if (isWorker) {
  scheduleFlush = useMessageChannel();
} else if (browserWindow === undefined && "function" === 'function') {
  scheduleFlush = attemptVertx();
} else {
  scheduleFlush = useSetTimeout();
}

function then(onFulfillment, onRejection) {
  var parent = this;

  var child = new this.constructor(noop);

  if (child[PROMISE_ID] === undefined) {
    makePromise(child);
  }

  var _state = parent._state;


  if (_state) {
    var callback = arguments[_state - 1];
    asap(function () {
      return invokeCallback(_state, child, callback, parent._result);
    });
  } else {
    subscribe(parent, child, onFulfillment, onRejection);
  }

  return child;
}

/**
  `Promise.resolve` returns a promise that will become resolved with the
  passed `value`. It is shorthand for the following:

  ```javascript
  let promise = new Promise(function(resolve, reject){
    resolve(1);
  });

  promise.then(function(value){
    // value === 1
  });
  ```

  Instead of writing the above, your code now simply becomes the following:

  ```javascript
  let promise = Promise.resolve(1);

  promise.then(function(value){
    // value === 1
  });
  ```

  @method resolve
  @static
  @param {Any} value value that the returned promise will be resolved with
  Useful for tooling.
  @return {Promise} a promise that will become fulfilled with the given
  `value`
*/
function resolve$1(object) {
  /*jshint validthis:true */
  var Constructor = this;

  if (object && typeof object === 'object' && object.constructor === Constructor) {
    return object;
  }

  var promise = new Constructor(noop);
  resolve(promise, object);
  return promise;
}

var PROMISE_ID = Math.random().toString(36).substring(2);

function noop() {}

var PENDING = void 0;
var FULFILLED = 1;
var REJECTED = 2;

function selfFulfillment() {
  return new TypeError("You cannot resolve a promise with itself");
}

function cannotReturnOwn() {
  return new TypeError('A promises callback cannot return that same promise.');
}

function tryThen(then$$1, value, fulfillmentHandler, rejectionHandler) {
  try {
    then$$1.call(value, fulfillmentHandler, rejectionHandler);
  } catch (e) {
    return e;
  }
}

function handleForeignThenable(promise, thenable, then$$1) {
  asap(function (promise) {
    var sealed = false;
    var error = tryThen(then$$1, thenable, function (value) {
      if (sealed) {
        return;
      }
      sealed = true;
      if (thenable !== value) {
        resolve(promise, value);
      } else {
        fulfill(promise, value);
      }
    }, function (reason) {
      if (sealed) {
        return;
      }
      sealed = true;

      reject(promise, reason);
    }, 'Settle: ' + (promise._label || ' unknown promise'));

    if (!sealed && error) {
      sealed = true;
      reject(promise, error);
    }
  }, promise);
}

function handleOwnThenable(promise, thenable) {
  if (thenable._state === FULFILLED) {
    fulfill(promise, thenable._result);
  } else if (thenable._state === REJECTED) {
    reject(promise, thenable._result);
  } else {
    subscribe(thenable, undefined, function (value) {
      return resolve(promise, value);
    }, function (reason) {
      return reject(promise, reason);
    });
  }
}

function handleMaybeThenable(promise, maybeThenable, then$$1) {
  if (maybeThenable.constructor === promise.constructor && then$$1 === then && maybeThenable.constructor.resolve === resolve$1) {
    handleOwnThenable(promise, maybeThenable);
  } else {
    if (then$$1 === undefined) {
      fulfill(promise, maybeThenable);
    } else if (isFunction(then$$1)) {
      handleForeignThenable(promise, maybeThenable, then$$1);
    } else {
      fulfill(promise, maybeThenable);
    }
  }
}

function resolve(promise, value) {
  if (promise === value) {
    reject(promise, selfFulfillment());
  } else if (objectOrFunction(value)) {
    var then$$1 = void 0;
    try {
      then$$1 = value.then;
    } catch (error) {
      reject(promise, error);
      return;
    }
    handleMaybeThenable(promise, value, then$$1);
  } else {
    fulfill(promise, value);
  }
}

function publishRejection(promise) {
  if (promise._onerror) {
    promise._onerror(promise._result);
  }

  publish(promise);
}

function fulfill(promise, value) {
  if (promise._state !== PENDING) {
    return;
  }

  promise._result = value;
  promise._state = FULFILLED;

  if (promise._subscribers.length !== 0) {
    asap(publish, promise);
  }
}

function reject(promise, reason) {
  if (promise._state !== PENDING) {
    return;
  }
  promise._state = REJECTED;
  promise._result = reason;

  asap(publishRejection, promise);
}

function subscribe(parent, child, onFulfillment, onRejection) {
  var _subscribers = parent._subscribers;
  var length = _subscribers.length;


  parent._onerror = null;

  _subscribers[length] = child;
  _subscribers[length + FULFILLED] = onFulfillment;
  _subscribers[length + REJECTED] = onRejection;

  if (length === 0 && parent._state) {
    asap(publish, parent);
  }
}

function publish(promise) {
  var subscribers = promise._subscribers;
  var settled = promise._state;

  if (subscribers.length === 0) {
    return;
  }

  var child = void 0,
      callback = void 0,
      detail = promise._result;

  for (var i = 0; i < subscribers.length; i += 3) {
    child = subscribers[i];
    callback = subscribers[i + settled];

    if (child) {
      invokeCallback(settled, child, callback, detail);
    } else {
      callback(detail);
    }
  }

  promise._subscribers.length = 0;
}

function invokeCallback(settled, promise, callback, detail) {
  var hasCallback = isFunction(callback),
      value = void 0,
      error = void 0,
      succeeded = true;

  if (hasCallback) {
    try {
      value = callback(detail);
    } catch (e) {
      succeeded = false;
      error = e;
    }

    if (promise === value) {
      reject(promise, cannotReturnOwn());
      return;
    }
  } else {
    value = detail;
  }

  if (promise._state !== PENDING) {
    // noop
  } else if (hasCallback && succeeded) {
    resolve(promise, value);
  } else if (succeeded === false) {
    reject(promise, error);
  } else if (settled === FULFILLED) {
    fulfill(promise, value);
  } else if (settled === REJECTED) {
    reject(promise, value);
  }
}

function initializePromise(promise, resolver) {
  try {
    resolver(function resolvePromise(value) {
      resolve(promise, value);
    }, function rejectPromise(reason) {
      reject(promise, reason);
    });
  } catch (e) {
    reject(promise, e);
  }
}

var id = 0;
function nextId() {
  return id++;
}

function makePromise(promise) {
  promise[PROMISE_ID] = id++;
  promise._state = undefined;
  promise._result = undefined;
  promise._subscribers = [];
}

function validationError() {
  return new Error('Array Methods must be provided an Array');
}

var Enumerator = function () {
  function Enumerator(Constructor, input) {
    this._instanceConstructor = Constructor;
    this.promise = new Constructor(noop);

    if (!this.promise[PROMISE_ID]) {
      makePromise(this.promise);
    }

    if (isArray(input)) {
      this.length = input.length;
      this._remaining = input.length;

      this._result = new Array(this.length);

      if (this.length === 0) {
        fulfill(this.promise, this._result);
      } else {
        this.length = this.length || 0;
        this._enumerate(input);
        if (this._remaining === 0) {
          fulfill(this.promise, this._result);
        }
      }
    } else {
      reject(this.promise, validationError());
    }
  }

  Enumerator.prototype._enumerate = function _enumerate(input) {
    for (var i = 0; this._state === PENDING && i < input.length; i++) {
      this._eachEntry(input[i], i);
    }
  };

  Enumerator.prototype._eachEntry = function _eachEntry(entry, i) {
    var c = this._instanceConstructor;
    var resolve$$1 = c.resolve;


    if (resolve$$1 === resolve$1) {
      var _then = void 0;
      var error = void 0;
      var didError = false;
      try {
        _then = entry.then;
      } catch (e) {
        didError = true;
        error = e;
      }

      if (_then === then && entry._state !== PENDING) {
        this._settledAt(entry._state, i, entry._result);
      } else if (typeof _then !== 'function') {
        this._remaining--;
        this._result[i] = entry;
      } else if (c === Promise$1) {
        var promise = new c(noop);
        if (didError) {
          reject(promise, error);
        } else {
          handleMaybeThenable(promise, entry, _then);
        }
        this._willSettleAt(promise, i);
      } else {
        this._willSettleAt(new c(function (resolve$$1) {
          return resolve$$1(entry);
        }), i);
      }
    } else {
      this._willSettleAt(resolve$$1(entry), i);
    }
  };

  Enumerator.prototype._settledAt = function _settledAt(state, i, value) {
    var promise = this.promise;


    if (promise._state === PENDING) {
      this._remaining--;

      if (state === REJECTED) {
        reject(promise, value);
      } else {
        this._result[i] = value;
      }
    }

    if (this._remaining === 0) {
      fulfill(promise, this._result);
    }
  };

  Enumerator.prototype._willSettleAt = function _willSettleAt(promise, i) {
    var enumerator = this;

    subscribe(promise, undefined, function (value) {
      return enumerator._settledAt(FULFILLED, i, value);
    }, function (reason) {
      return enumerator._settledAt(REJECTED, i, reason);
    });
  };

  return Enumerator;
}();

/**
  `Promise.all` accepts an array of promises, and returns a new promise which
  is fulfilled with an array of fulfillment values for the passed promises, or
  rejected with the reason of the first passed promise to be rejected. It casts all
  elements of the passed iterable to promises as it runs this algorithm.

  Example:

  ```javascript
  let promise1 = resolve(1);
  let promise2 = resolve(2);
  let promise3 = resolve(3);
  let promises = [ promise1, promise2, promise3 ];

  Promise.all(promises).then(function(array){
    // The array here would be [ 1, 2, 3 ];
  });
  ```

  If any of the `promises` given to `all` are rejected, the first promise
  that is rejected will be given as an argument to the returned promises's
  rejection handler. For example:

  Example:

  ```javascript
  let promise1 = resolve(1);
  let promise2 = reject(new Error("2"));
  let promise3 = reject(new Error("3"));
  let promises = [ promise1, promise2, promise3 ];

  Promise.all(promises).then(function(array){
    // Code here never runs because there are rejected promises!
  }, function(error) {
    // error.message === "2"
  });
  ```

  @method all
  @static
  @param {Array} entries array of promises
  @param {String} label optional string for labeling the promise.
  Useful for tooling.
  @return {Promise} promise that is fulfilled when all `promises` have been
  fulfilled, or rejected if any of them become rejected.
  @static
*/
function all(entries) {
  return new Enumerator(this, entries).promise;
}

/**
  `Promise.race` returns a new promise which is settled in the same way as the
  first passed promise to settle.

  Example:

  ```javascript
  let promise1 = new Promise(function(resolve, reject){
    setTimeout(function(){
      resolve('promise 1');
    }, 200);
  });

  let promise2 = new Promise(function(resolve, reject){
    setTimeout(function(){
      resolve('promise 2');
    }, 100);
  });

  Promise.race([promise1, promise2]).then(function(result){
    // result === 'promise 2' because it was resolved before promise1
    // was resolved.
  });
  ```

  `Promise.race` is deterministic in that only the state of the first
  settled promise matters. For example, even if other promises given to the
  `promises` array argument are resolved, but the first settled promise has
  become rejected before the other promises became fulfilled, the returned
  promise will become rejected:

  ```javascript
  let promise1 = new Promise(function(resolve, reject){
    setTimeout(function(){
      resolve('promise 1');
    }, 200);
  });

  let promise2 = new Promise(function(resolve, reject){
    setTimeout(function(){
      reject(new Error('promise 2'));
    }, 100);
  });

  Promise.race([promise1, promise2]).then(function(result){
    // Code here never runs
  }, function(reason){
    // reason.message === 'promise 2' because promise 2 became rejected before
    // promise 1 became fulfilled
  });
  ```

  An example real-world use case is implementing timeouts:

  ```javascript
  Promise.race([ajax('foo.json'), timeout(5000)])
  ```

  @method race
  @static
  @param {Array} promises array of promises to observe
  Useful for tooling.
  @return {Promise} a promise which settles in the same way as the first passed
  promise to settle.
*/
function race(entries) {
  /*jshint validthis:true */
  var Constructor = this;

  if (!isArray(entries)) {
    return new Constructor(function (_, reject) {
      return reject(new TypeError('You must pass an array to race.'));
    });
  } else {
    return new Constructor(function (resolve, reject) {
      var length = entries.length;
      for (var i = 0; i < length; i++) {
        Constructor.resolve(entries[i]).then(resolve, reject);
      }
    });
  }
}

/**
  `Promise.reject` returns a promise rejected with the passed `reason`.
  It is shorthand for the following:

  ```javascript
  let promise = new Promise(function(resolve, reject){
    reject(new Error('WHOOPS'));
  });

  promise.then(function(value){
    // Code here doesn't run because the promise is rejected!
  }, function(reason){
    // reason.message === 'WHOOPS'
  });
  ```

  Instead of writing the above, your code now simply becomes the following:

  ```javascript
  let promise = Promise.reject(new Error('WHOOPS'));

  promise.then(function(value){
    // Code here doesn't run because the promise is rejected!
  }, function(reason){
    // reason.message === 'WHOOPS'
  });
  ```

  @method reject
  @static
  @param {Any} reason value that the returned promise will be rejected with.
  Useful for tooling.
  @return {Promise} a promise rejected with the given `reason`.
*/
function reject$1(reason) {
  /*jshint validthis:true */
  var Constructor = this;
  var promise = new Constructor(noop);
  reject(promise, reason);
  return promise;
}

function needsResolver() {
  throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
}

function needsNew() {
  throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
}

/**
  Promise objects represent the eventual result of an asynchronous operation. The
  primary way of interacting with a promise is through its `then` method, which
  registers callbacks to receive either a promise's eventual value or the reason
  why the promise cannot be fulfilled.

  Terminology
  -----------

  - `promise` is an object or function with a `then` method whose behavior conforms to this specification.
  - `thenable` is an object or function that defines a `then` method.
  - `value` is any legal JavaScript value (including undefined, a thenable, or a promise).
  - `exception` is a value that is thrown using the throw statement.
  - `reason` is a value that indicates why a promise was rejected.
  - `settled` the final resting state of a promise, fulfilled or rejected.

  A promise can be in one of three states: pending, fulfilled, or rejected.

  Promises that are fulfilled have a fulfillment value and are in the fulfilled
  state.  Promises that are rejected have a rejection reason and are in the
  rejected state.  A fulfillment value is never a thenable.

  Promises can also be said to *resolve* a value.  If this value is also a
  promise, then the original promise's settled state will match the value's
  settled state.  So a promise that *resolves* a promise that rejects will
  itself reject, and a promise that *resolves* a promise that fulfills will
  itself fulfill.


  Basic Usage:
  ------------

  ```js
  let promise = new Promise(function(resolve, reject) {
    // on success
    resolve(value);

    // on failure
    reject(reason);
  });

  promise.then(function(value) {
    // on fulfillment
  }, function(reason) {
    // on rejection
  });
  ```

  Advanced Usage:
  ---------------

  Promises shine when abstracting away asynchronous interactions such as
  `XMLHttpRequest`s.

  ```js
  function getJSON(url) {
    return new Promise(function(resolve, reject){
      let xhr = new XMLHttpRequest();

      xhr.open('GET', url);
      xhr.onreadystatechange = handler;
      xhr.responseType = 'json';
      xhr.setRequestHeader('Accept', 'application/json');
      xhr.send();

      function handler() {
        if (this.readyState === this.DONE) {
          if (this.status === 200) {
            resolve(this.response);
          } else {
            reject(new Error('getJSON: `' + url + '` failed with status: [' + this.status + ']'));
          }
        }
      };
    });
  }

  getJSON('/posts.json').then(function(json) {
    // on fulfillment
  }, function(reason) {
    // on rejection
  });
  ```

  Unlike callbacks, promises are great composable primitives.

  ```js
  Promise.all([
    getJSON('/posts'),
    getJSON('/comments')
  ]).then(function(values){
    values[0] // => postsJSON
    values[1] // => commentsJSON

    return values;
  });
  ```

  @class Promise
  @param {Function} resolver
  Useful for tooling.
  @constructor
*/

var Promise$1 = function () {
  function Promise(resolver) {
    this[PROMISE_ID] = nextId();
    this._result = this._state = undefined;
    this._subscribers = [];

    if (noop !== resolver) {
      typeof resolver !== 'function' && needsResolver();
      this instanceof Promise ? initializePromise(this, resolver) : needsNew();
    }
  }

  /**
  The primary way of interacting with a promise is through its `then` method,
  which registers callbacks to receive either a promise's eventual value or the
  reason why the promise cannot be fulfilled.
   ```js
  findUser().then(function(user){
    // user is available
  }, function(reason){
    // user is unavailable, and you are given the reason why
  });
  ```
   Chaining
  --------
   The return value of `then` is itself a promise.  This second, 'downstream'
  promise is resolved with the return value of the first promise's fulfillment
  or rejection handler, or rejected if the handler throws an exception.
   ```js
  findUser().then(function (user) {
    return user.name;
  }, function (reason) {
    return 'default name';
  }).then(function (userName) {
    // If `findUser` fulfilled, `userName` will be the user's name, otherwise it
    // will be `'default name'`
  });
   findUser().then(function (user) {
    throw new Error('Found user, but still unhappy');
  }, function (reason) {
    throw new Error('`findUser` rejected and we're unhappy');
  }).then(function (value) {
    // never reached
  }, function (reason) {
    // if `findUser` fulfilled, `reason` will be 'Found user, but still unhappy'.
    // If `findUser` rejected, `reason` will be '`findUser` rejected and we're unhappy'.
  });
  ```
  If the downstream promise does not specify a rejection handler, rejection reasons will be propagated further downstream.
   ```js
  findUser().then(function (user) {
    throw new PedagogicalException('Upstream error');
  }).then(function (value) {
    // never reached
  }).then(function (value) {
    // never reached
  }, function (reason) {
    // The `PedgagocialException` is propagated all the way down to here
  });
  ```
   Assimilation
  ------------
   Sometimes the value you want to propagate to a downstream promise can only be
  retrieved asynchronously. This can be achieved by returning a promise in the
  fulfillment or rejection handler. The downstream promise will then be pending
  until the returned promise is settled. This is called *assimilation*.
   ```js
  findUser().then(function (user) {
    return findCommentsByAuthor(user);
  }).then(function (comments) {
    // The user's comments are now available
  });
  ```
   If the assimliated promise rejects, then the downstream promise will also reject.
   ```js
  findUser().then(function (user) {
    return findCommentsByAuthor(user);
  }).then(function (comments) {
    // If `findCommentsByAuthor` fulfills, we'll have the value here
  }, function (reason) {
    // If `findCommentsByAuthor` rejects, we'll have the reason here
  });
  ```
   Simple Example
  --------------
   Synchronous Example
   ```javascript
  let result;
   try {
    result = findResult();
    // success
  } catch(reason) {
    // failure
  }
  ```
   Errback Example
   ```js
  findResult(function(result, err){
    if (err) {
      // failure
    } else {
      // success
    }
  });
  ```
   Promise Example;
   ```javascript
  findResult().then(function(result){
    // success
  }, function(reason){
    // failure
  });
  ```
   Advanced Example
  --------------
   Synchronous Example
   ```javascript
  let author, books;
   try {
    author = findAuthor();
    books  = findBooksByAuthor(author);
    // success
  } catch(reason) {
    // failure
  }
  ```
   Errback Example
   ```js
   function foundBooks(books) {
   }
   function failure(reason) {
   }
   findAuthor(function(author, err){
    if (err) {
      failure(err);
      // failure
    } else {
      try {
        findBoooksByAuthor(author, function(books, err) {
          if (err) {
            failure(err);
          } else {
            try {
              foundBooks(books);
            } catch(reason) {
              failure(reason);
            }
          }
        });
      } catch(error) {
        failure(err);
      }
      // success
    }
  });
  ```
   Promise Example;
   ```javascript
  findAuthor().
    then(findBooksByAuthor).
    then(function(books){
      // found books
  }).catch(function(reason){
    // something went wrong
  });
  ```
   @method then
  @param {Function} onFulfilled
  @param {Function} onRejected
  Useful for tooling.
  @return {Promise}
  */

  /**
  `catch` is simply sugar for `then(undefined, onRejection)` which makes it the same
  as the catch block of a try/catch statement.
  ```js
  function findAuthor(){
  throw new Error('couldn't find that author');
  }
  // synchronous
  try {
  findAuthor();
  } catch(reason) {
  // something went wrong
  }
  // async with promises
  findAuthor().catch(function(reason){
  // something went wrong
  });
  ```
  @method catch
  @param {Function} onRejection
  Useful for tooling.
  @return {Promise}
  */


  Promise.prototype.catch = function _catch(onRejection) {
    return this.then(null, onRejection);
  };

  /**
    `finally` will be invoked regardless of the promise's fate just as native
    try/catch/finally behaves
  
    Synchronous example:
  
    ```js
    findAuthor() {
      if (Math.random() > 0.5) {
        throw new Error();
      }
      return new Author();
    }
  
    try {
      return findAuthor(); // succeed or fail
    } catch(error) {
      return findOtherAuther();
    } finally {
      // always runs
      // doesn't affect the return value
    }
    ```
  
    Asynchronous example:
  
    ```js
    findAuthor().catch(function(reason){
      return findOtherAuther();
    }).finally(function(){
      // author was either found, or not
    });
    ```
  
    @method finally
    @param {Function} callback
    @return {Promise}
  */


  Promise.prototype.finally = function _finally(callback) {
    var promise = this;
    var constructor = promise.constructor;

    if (isFunction(callback)) {
      return promise.then(function (value) {
        return constructor.resolve(callback()).then(function () {
          return value;
        });
      }, function (reason) {
        return constructor.resolve(callback()).then(function () {
          throw reason;
        });
      });
    }

    return promise.then(callback, callback);
  };

  return Promise;
}();

Promise$1.prototype.then = then;
Promise$1.all = all;
Promise$1.race = race;
Promise$1.resolve = resolve$1;
Promise$1.reject = reject$1;
Promise$1._setScheduler = setScheduler;
Promise$1._setAsap = setAsap;
Promise$1._asap = asap;

/*global self*/
function polyfill() {
  var local = void 0;

  if (typeof __webpack_require__.g !== 'undefined') {
    local = __webpack_require__.g;
  } else if (typeof self !== 'undefined') {
    local = self;
  } else {
    try {
      local = Function('return this')();
    } catch (e) {
      throw new Error('polyfill failed because global object is unavailable in this environment');
    }
  }

  var P = local.Promise;

  if (P) {
    var promiseToString = null;
    try {
      promiseToString = Object.prototype.toString.call(P.resolve());
    } catch (e) {
      // silently ignored
    }

    if (promiseToString === '[object Promise]' && !P.cast) {
      return;
    }
  }

  local.Promise = Promise$1;
}

// Strange compat..
Promise$1.polyfill = polyfill;
Promise$1.Promise = Promise$1;

return Promise$1;

})));



//# sourceMappingURL=es6-promise.map


/***/ }),

/***/ "./src/commands/onMessageSendHandler.ts":
/*!**********************************************!*\
  !*** ./src/commands/onMessageSendHandler.ts ***!
  \**********************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   onMessageSendHandler: function() { return /* binding */ onMessageSendHandler; }
/* harmony export */ });
/* harmony import */ var _services_auth__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../services/auth */ "./src/services/auth.ts");
/* harmony import */ var _utils_storage__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../utils/storage */ "./src/utils/storage.ts");
/* harmony import */ var _utils_constants__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utils/constants */ "./src/utils/constants.ts");
/* harmony import */ var _services_singlecaseDocuments__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../services/singlecaseDocuments */ "./src/services/singlecaseDocuments.ts");
/* harmony import */ var _utils_filedCache__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/filedCache */ "./src/utils/filedCache.ts");
/* provided dependency */ var Promise = __webpack_require__(/*! es6-promise */ "./node_modules/es6-promise/dist/es6-promise.js")["Promise"];
/* global Office, OfficeRuntime */
var __awaiter = undefined && undefined.__awaiter || function (thisArg, _arguments, P, generator) {
  function adopt(value) {
    return value instanceof P ? value : new P(function (resolve) {
      resolve(value);
    });
  }
  return new (P || (P = Promise))(function (resolve, reject) {
    function fulfilled(value) {
      try {
        step(generator.next(value));
      } catch (e) {
        reject(e);
      }
    }
    function rejected(value) {
      try {
        step(generator["throw"](value));
      } catch (e) {
        reject(e);
      }
    }
    function step(result) {
      result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected);
    }
    step((generator = generator.apply(thisArg, _arguments || [])).next());
  });
};
var __generator = undefined && undefined.__generator || function (thisArg, body) {
  var _ = {
      label: 0,
      sent: function sent() {
        if (t[0] & 1) throw t[1];
        return t[1];
      },
      trys: [],
      ops: []
    },
    f,
    y,
    t,
    g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
  return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function () {
    return this;
  }), g;
  function verb(n) {
    return function (v) {
      return step([n, v]);
    };
  }
  function step(op) {
    if (f) throw new TypeError("Generator is already executing.");
    while (g && (g = 0, op[0] && (_ = 0)), _) try {
      if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
      if (y = 0, t) op = [op[0] & 2, t.value];
      switch (op[0]) {
        case 0:
        case 1:
          t = op;
          break;
        case 4:
          _.label++;
          return {
            value: op[1],
            done: false
          };
        case 5:
          _.label++;
          y = op[1];
          op = [0];
          continue;
        case 7:
          op = _.ops.pop();
          _.trys.pop();
          continue;
        default:
          if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) {
            _ = 0;
            continue;
          }
          if (op[0] === 3 && (!t || op[1] > t[0] && op[1] < t[3])) {
            _.label = op[1];
            break;
          }
          if (op[0] === 6 && _.label < t[1]) {
            _.label = t[1];
            t = op;
            break;
          }
          if (t && _.label < t[2]) {
            _.label = t[2];
            _.ops.push(op);
            break;
          }
          if (t[2]) _.ops.pop();
          _.trys.pop();
          continue;
      }
      op = body.call(thisArg, _);
    } catch (e) {
      op = [6, e];
      y = 0;
    } finally {
      f = t = 0;
    }
    if (op[0] & 5) throw op[1];
    return {
      value: op[0] ? op[1] : void 0,
      done: true
    };
  }
};





var T_ITEMKEY_MS = 2000;
var T_STORAGE_MS = 1500;
var T_FETCH_MS = 10000;
var T_SUBJECT_MS = 1500;
var T_BODY_MS = 2500;
var CONV_CTX_KEY_PREFIX = "sc_conv_ctx:";
var LAST_FILED_CTX_KEY = "sc_last_filed_ctx";
function withTimeout(p, ms) {
  return new Promise(function (resolve, reject) {
    var t = setTimeout(function () {
      return reject(new Error("timeout"));
    }, ms);
    p.then(function (v) {
      clearTimeout(t);
      resolve(v);
    }, function (e) {
      clearTimeout(t);
      reject(e);
    });
  });
}
function normalizeHost(host) {
  var v = (host || "").trim().toLowerCase();
  if (!v) return "";
  return v.replace(/^https?:\/\//i, "").split("/")[0];
}
function safeFileName(value) {
  var v = (value || "").trim();
  var cleaned = v.replace(/[<>:"/\\|?*\x00-\x1F]/g, " ").replace(/\s+/g, " ").trim();
  return cleaned.slice(0, 80) || "email";
}
function toBase64Utf8(text) {
  var bytes = new TextEncoder().encode(text);
  var binary = "";
  for (var i = 0; i < bytes.length; i += 1) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}
function getConversationIdSafe() {
  try {
    var item = Office.context.mailbox.item;
    return String((item === null || item === void 0 ? void 0 : item.conversationId) || (item === null || item === void 0 ? void 0 : item.conversationKey) || "").trim();
  } catch (_a) {
    return "";
  }
}
function persistFiledCtx(caseId, emailDocId) {
  return __awaiter(this, void 0, void 0, function () {
    var cid, did, payload, _a, convId, _b;
    return __generator(this, function (_c) {
      switch (_c.label) {
        case 0:
          cid = String(caseId || "").trim();
          did = String(emailDocId || "").trim();
          if (!cid || !did) return [2 /*return*/];
          payload = JSON.stringify({
            caseId: cid,
            emailDocId: did
          });
          _c.label = 1;
        case 1:
          _c.trys.push([1, 3,, 4]);
          return [4 /*yield*/, (0,_utils_storage__WEBPACK_IMPORTED_MODULE_1__.setStored)(LAST_FILED_CTX_KEY, payload)];
        case 2:
          _c.sent();
          return [3 /*break*/, 4];
        case 3:
          _a = _c.sent();
          return [3 /*break*/, 4];
        case 4:
          _c.trys.push([4, 7,, 8]);
          convId = getConversationIdSafe();
          if (!convId) return [3 /*break*/, 6];
          return [4 /*yield*/, (0,_utils_storage__WEBPACK_IMPORTED_MODULE_1__.setStored)("".concat(CONV_CTX_KEY_PREFIX).concat(convId), payload)];
        case 5:
          _c.sent();
          _c.label = 6;
        case 6:
          return [3 /*break*/, 8];
        case 7:
          _b = _c.sent();
          return [3 /*break*/, 8];
        case 8:
          return [2 /*return*/];
      }
    });
  });
}
function getCandidateItemKeysRuntime() {
  return __awaiter(this, void 0, void 0, function () {
    var item, keys, direct, asyncId, e_1, conv, created;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          item = Office.context.mailbox.item;
          if (!item) {
            console.warn("[getCandidateItemKeysRuntime] No item available");
            return [2 /*return*/, []];
          }
          console.log("[getCandidateItemKeysRuntime] Item properties:", {
            hasItemId: !!item.itemId,
            itemId: String(item.itemId || "").substring(0, 20),
            hasConversationId: !!item.conversationId,
            conversationId: String(item.conversationId || "").substring(0, 20),
            hasConversationKey: !!item.conversationKey,
            hasDateTimeCreated: !!item.dateTimeCreated,
            hasGetItemIdAsync: typeof item.getItemIdAsync === "function",
            itemType: item.itemType
          });
          keys = [];
          direct = String(item.itemId || "").trim();
          if (direct) keys.push(direct);
          if (!(typeof item.getItemIdAsync === "function")) return [3 /*break*/, 4];
          _a.label = 1;
        case 1:
          _a.trys.push([1, 3,, 4]);
          return [4 /*yield*/, new Promise(function (resolve) {
            item.getItemIdAsync(function (res) {
              if ((res === null || res === void 0 ? void 0 : res.status) === Office.AsyncResultStatus.Succeeded) resolve(String(res.value || ""));else resolve("");
            });
          })];
        case 2:
          asyncId = _a.sent();
          if (asyncId) {
            console.log("[getCandidateItemKeysRuntime] getItemIdAsync returned:", asyncId.substring(0, 20));
            keys.push(asyncId);
          }
          return [3 /*break*/, 4];
        case 3:
          e_1 = _a.sent();
          console.warn("[getCandidateItemKeysRuntime] getItemIdAsync failed:", e_1);
          return [3 /*break*/, 4];
        case 4:
          conv = String(item.conversationId || item.conversationKey || "").trim();
          if (conv) keys.push("draft:".concat(conv));
          created = String(item.dateTimeCreated || "").trim();
          if (created) keys.push("draft:".concat(created));
          // Always include fallback keys for new compose emails
          keys.push("draft:current");
          keys.push("last_compose");
          console.log("[getCandidateItemKeysRuntime] Generated keys:", keys);
          return [2 /*return*/, Array.from(new Set(keys.filter(Boolean)))];
      }
    });
  });
}
function readIntentAny(itemKeys) {
  return __awaiter(this, void 0, void 0, function () {
    var _i, itemKeys_1, k, key, raw, e_2, obj, caseId, autoFileOnSend, baseCaseId, baseEmailDocId, e_3;
    var _a;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          _i = 0, itemKeys_1 = itemKeys;
          _b.label = 1;
        case 1:
          if (!(_i < itemKeys_1.length)) return [3 /*break*/, 9];
          k = itemKeys_1[_i];
          key = "sc_intent:".concat(k);
          console.log("[readIntentAny] Trying key:", key);
          _b.label = 2;
        case 2:
          _b.trys.push([2, 7,, 8]);
          raw = null;
          if (!(typeof OfficeRuntime !== "undefined" && (OfficeRuntime === null || OfficeRuntime === void 0 ? void 0 : OfficeRuntime.storage))) return [3 /*break*/, 6];
          _b.label = 3;
        case 3:
          _b.trys.push([3, 5,, 6]);
          return [4 /*yield*/, OfficeRuntime.storage.getItem(key)];
        case 4:
          raw = _b.sent();
          if (raw) console.log("[readIntentAny] Found in OfficeRuntime.storage");
          return [3 /*break*/, 6];
        case 5:
          e_2 = _b.sent();
          console.warn("[readIntentAny] OfficeRuntime.storage.getItem failed:", e_2);
          return [3 /*break*/, 6];
        case 6:
          if (!raw && ((_a = Office === null || Office === void 0 ? void 0 : Office.context) === null || _a === void 0 ? void 0 : _a.roamingSettings)) {
            try {
              raw = Office.context.roamingSettings.get(key);
              if (raw) console.log("[readIntentAny] Found in roamingSettings");
            } catch (e) {
              console.warn("[readIntentAny] roamingSettings.get failed:", e);
            }
          }
          if (!raw) return [3 /*break*/, 8];
          obj = JSON.parse(String(raw));
          caseId = String((obj === null || obj === void 0 ? void 0 : obj.caseId) || "").trim();
          autoFileOnSend = Boolean(obj === null || obj === void 0 ? void 0 : obj.autoFileOnSend);
          baseCaseId = String((obj === null || obj === void 0 ? void 0 : obj.baseCaseId) || "").trim();
          baseEmailDocId = String((obj === null || obj === void 0 ? void 0 : obj.baseEmailDocId) || "").trim();
          if (!caseId) return [3 /*break*/, 8];
          console.log("[readIntentAny] Intent found:", {
            itemKey: k,
            caseId: caseId,
            autoFileOnSend: autoFileOnSend,
            hasBase: !!(baseCaseId && baseEmailDocId)
          });
          return [2 /*return*/, {
            itemKey: k,
            caseId: caseId,
            autoFileOnSend: autoFileOnSend,
            baseCaseId: baseCaseId || undefined,
            baseEmailDocId: baseEmailDocId || undefined
          }];
        case 7:
          e_3 = _b.sent();
          console.warn("[readIntentAny] Failed to read intent for key:", key, e_3);
          return [3 /*break*/, 8];
        case 8:
          _i++;
          return [3 /*break*/, 1];
        case 9:
          console.warn("[readIntentAny] No intent found for any key");
          return [2 /*return*/, null];
      }
    });
  });
}
function getSubjectRuntime() {
  return __awaiter(this, void 0, void 0, function () {
    var item, subj, v;
    var _a;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          item = Office.context.mailbox.item;
          if (!item) {
            console.warn("[getSubjectRuntime] No item available");
            return [2 /*return*/, ""];
          }
          console.log("[getSubjectRuntime] Item type:", item.itemType, "Mode:", item.itemClass);
          if (typeof item.subject === "string") {
            subj = String(item.subject || "");
            console.log("[getSubjectRuntime] Direct string subject:", subj);
            return [2 /*return*/, subj];
          }
          if (!((_a = item === null || item === void 0 ? void 0 : item.subject) === null || _a === void 0 ? void 0 : _a.getAsync)) return [3 /*break*/, 2];
          return [4 /*yield*/, new Promise(function (resolve) {
            item.subject.getAsync(function (res) {
              if ((res === null || res === void 0 ? void 0 : res.status) === Office.AsyncResultStatus.Succeeded) {
                var subj = String(res.value || "");
                console.log("[getSubjectRuntime] Async subject:", subj);
                resolve(subj);
              } else {
                console.warn("[getSubjectRuntime] getAsync failed:", res === null || res === void 0 ? void 0 : res.error);
                resolve("");
              }
            });
          })];
        case 1:
          v = _b.sent();
          return [2 /*return*/, v || ""];
        case 2:
          console.warn("[getSubjectRuntime] No subject API available");
          return [2 /*return*/, ""];
      }
    });
  });
}
function getBodyTextRuntime() {
  return __awaiter(this, void 0, void 0, function () {
    var item, text;
    var _a;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          item = Office.context.mailbox.item;
          if (!((_a = item === null || item === void 0 ? void 0 : item.body) === null || _a === void 0 ? void 0 : _a.getAsync)) return [2 /*return*/, ""];
          return [4 /*yield*/, new Promise(function (resolve) {
            item.body.getAsync(Office.CoercionType.Text, function (res) {
              if ((res === null || res === void 0 ? void 0 : res.status) === Office.AsyncResultStatus.Succeeded) resolve(String(res.value || ""));else resolve("");
            });
          })];
        case 1:
          text = _b.sent();
          return [2 /*return*/, String(text || "")];
      }
    });
  });
}
function showInfo(message) {
  return __awaiter(this, void 0, void 0, function () {
    var item_1, _a;
    var _b;
    return __generator(this, function (_c) {
      switch (_c.label) {
        case 0:
          _c.trys.push([0, 2,, 3]);
          item_1 = Office.context.mailbox.item;
          if (!((_b = item_1 === null || item_1 === void 0 ? void 0 : item_1.notificationMessages) === null || _b === void 0 ? void 0 : _b.replaceAsync)) return [2 /*return*/];
          return [4 /*yield*/, new Promise(function (resolve) {
            item_1.notificationMessages.replaceAsync("sc_send", {
              type: "informationalMessage",
              message: message,
              icon: "Icon.16x16",
              persistent: false
            }, function () {
              return resolve();
            });
          })];
        case 1:
          _c.sent();
          return [3 /*break*/, 3];
        case 2:
          _a = _c.sent();
          return [3 /*break*/, 3];
        case 3:
          return [2 /*return*/];
      }
    });
  });
}
function onMessageSendHandler(event) {
  return __awaiter(this, void 0, void 0, function () {
    var done, finish, keys, intent, isFallbackKey, realItemId, intentValue, realKey, fallbackKey, e_4, token, hostRaw, host, subject, bodyText, itemFrom, fromEmail, fromName, conversationId, baseName, emailText, emailBase64, existingDoc, e_5, shouldUploadVersion, conversationId_1, created, docs, createdDocId, conversationId_2, e_6, msg, errorHint, _a;
    var _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q, _r, _s;
    return __generator(this, function (_t) {
      switch (_t.label) {
        case 0:
          console.log("[onMessageSendHandler] Handler fired");
          console.log("[onMessageSendHandler] Platform info", {
            hasOfficeRuntime: typeof OfficeRuntime !== "undefined",
            hasOfficeRuntimeStorage: typeof (OfficeRuntime === null || OfficeRuntime === void 0 ? void 0 : OfficeRuntime.storage) !== "undefined",
            hasRoamingSettings: !!((_b = Office === null || Office === void 0 ? void 0 : Office.context) === null || _b === void 0 ? void 0 : _b.roamingSettings),
            host: (_e = (_d = (_c = Office === null || Office === void 0 ? void 0 : Office.context) === null || _c === void 0 ? void 0 : _c.mailbox) === null || _d === void 0 ? void 0 : _d.diagnostics) === null || _e === void 0 ? void 0 : _e.hostName,
            hostVersion: (_h = (_g = (_f = Office === null || Office === void 0 ? void 0 : Office.context) === null || _f === void 0 ? void 0 : _f.mailbox) === null || _g === void 0 ? void 0 : _g.diagnostics) === null || _h === void 0 ? void 0 : _h.hostVersion
          });
          done = false;
          finish = function finish(allowEvent, errorMessage) {
            if (done) return;
            done = true;
            console.log("[onMessageSendHandler] Finishing", {
              allowEvent: allowEvent,
              hasErrorMessage: !!errorMessage
            });
            try {
              if (errorMessage) event.completed({
                allowEvent: allowEvent,
                errorMessage: errorMessage
              });else event.completed({
                allowEvent: allowEvent
              });
            } catch (e) {
              console.error("[onMessageSendHandler] Error in event.completed:", e);
            }
          };
          _t.label = 1;
        case 1:
          _t.trys.push([1, 42,, 47]);
          console.log("[onMessageSendHandler] Clearing expired auth");
          return [4 /*yield*/, withTimeout((0,_services_auth__WEBPACK_IMPORTED_MODULE_0__.clearAuthIfExpiredRuntime)(), 700)];
        case 2:
          _t.sent();
          console.log("[onMessageSendHandler] Getting candidate item keys");
          return [4 /*yield*/, withTimeout(getCandidateItemKeysRuntime(), T_ITEMKEY_MS)];
        case 3:
          keys = _t.sent();
          console.log("[onMessageSendHandler] Item keys:", keys);
          if (keys.length === 0) {
            console.log("[onMessageSendHandler] No item keys found, skipping");
            finish(true);
            return [2 /*return*/];
          }
          console.log("[onMessageSendHandler] Reading intent from storage", {
            storageType: typeof OfficeRuntime !== "undefined" && (OfficeRuntime === null || OfficeRuntime === void 0 ? void 0 : OfficeRuntime.storage) ? "OfficeRuntime" : "roamingSettings",
            keysToTry: keys
          });
          return [4 /*yield*/, withTimeout(readIntentAny(keys), T_STORAGE_MS)];
        case 4:
          intent = _t.sent();
          console.log("[onMessageSendHandler] Intent:", intent, {
            found: !!intent,
            foundUnderKey: intent === null || intent === void 0 ? void 0 : intent.itemKey
          });
          if (!(intent === null || intent === void 0 ? void 0 : intent.autoFileOnSend) || !intent.caseId) {
            console.log("[onMessageSendHandler] No auto-file intent or case ID, skipping");
            finish(true);
            return [2 /*return*/];
          }
          _t.label = 5;
        case 5:
          _t.trys.push([5, 14,, 15]);
          isFallbackKey = intent.itemKey === "draft:current" || intent.itemKey === "last_compose";
          if (!isFallbackKey) return [3 /*break*/, 13];
          realItemId = keys.find(function (k) {
            return !k.startsWith("draft:") && k !== "last_compose";
          });
          if (!realItemId) return [3 /*break*/, 13];
          console.log("[onMessageSendHandler] Migrating intent from fallback", {
            from: intent.itemKey,
            to: realItemId
          });
          intentValue = JSON.stringify({
            caseId: intent.caseId,
            autoFileOnSend: intent.autoFileOnSend,
            baseCaseId: intent.baseCaseId || "",
            baseEmailDocId: intent.baseEmailDocId || ""
          });
          realKey = "sc_intent:".concat(realItemId);
          if (!(typeof OfficeRuntime !== "undefined" && (OfficeRuntime === null || OfficeRuntime === void 0 ? void 0 : OfficeRuntime.storage))) return [3 /*break*/, 7];
          return [4 /*yield*/, OfficeRuntime.storage.setItem(realKey, intentValue)];
        case 6:
          _t.sent();
          console.log("[onMessageSendHandler] Migrated to OfficeRuntime.storage");
          return [3 /*break*/, 9];
        case 7:
          if (!((_j = Office === null || Office === void 0 ? void 0 : Office.context) === null || _j === void 0 ? void 0 : _j.roamingSettings)) return [3 /*break*/, 9];
          Office.context.roamingSettings.set(realKey, intentValue);
          return [4 /*yield*/, new Promise(function (resolve) {
            Office.context.roamingSettings.saveAsync(function () {
              return resolve();
            });
          })];
        case 8:
          _t.sent();
          console.log("[onMessageSendHandler] Migrated to roamingSettings");
          _t.label = 9;
        case 9:
          fallbackKey = "sc_intent:".concat(intent.itemKey);
          if (!(typeof OfficeRuntime !== "undefined" && (OfficeRuntime === null || OfficeRuntime === void 0 ? void 0 : OfficeRuntime.storage))) return [3 /*break*/, 11];
          return [4 /*yield*/, OfficeRuntime.storage.removeItem(fallbackKey)];
        case 10:
          _t.sent();
          console.log("[onMessageSendHandler] Cleared fallback key from OfficeRuntime.storage");
          return [3 /*break*/, 13];
        case 11:
          if (!((_k = Office === null || Office === void 0 ? void 0 : Office.context) === null || _k === void 0 ? void 0 : _k.roamingSettings)) return [3 /*break*/, 13];
          Office.context.roamingSettings.remove(fallbackKey);
          return [4 /*yield*/, new Promise(function (resolve) {
            Office.context.roamingSettings.saveAsync(function () {
              return resolve();
            });
          })];
        case 12:
          _t.sent();
          console.log("[onMessageSendHandler] Cleared fallback key from roamingSettings");
          _t.label = 13;
        case 13:
          return [3 /*break*/, 15];
        case 14:
          e_4 = _t.sent();
          console.warn("[onMessageSendHandler] Intent migration failed (non-critical):", e_4);
          return [3 /*break*/, 15];
        case 15:
          console.log("[onMessageSendHandler] Getting auth token");
          return [4 /*yield*/, withTimeout((0,_services_auth__WEBPACK_IMPORTED_MODULE_0__.getAuthRuntime)(), 900)];
        case 16:
          token = _t.sent().token;
          if (!!token) return [3 /*break*/, 18];
          console.error("[onMessageSendHandler] No auth token available");
          return [4 /*yield*/, showInfo("SingleCase: chybí přihlášení, nelze zařadit při odeslání.")];
        case 17:
          _t.sent();
          finish(true);
          return [2 /*return*/];
        case 18:
          console.log("[onMessageSendHandler] Token retrieved", {
            tokenPrefix: token.slice(0, 10)
          });
          console.log("[onMessageSendHandler] Getting workspace host");
          return [4 /*yield*/, (0,_utils_storage__WEBPACK_IMPORTED_MODULE_1__.getStored)(_utils_constants__WEBPACK_IMPORTED_MODULE_2__.STORAGE_KEYS.workspaceHost)];
        case 19:
          hostRaw = _t.sent() || "";
          host = normalizeHost(hostRaw);
          console.log("[onMessageSendHandler] Workspace host", {
            hostRaw: hostRaw,
            normalized: host
          });
          if (!!host) return [3 /*break*/, 21];
          console.error("[onMessageSendHandler] No workspace host configured");
          return [4 /*yield*/, showInfo("SingleCase: chybí workspace URL, nelze zařadit při odeslání.")];
        case 20:
          _t.sent();
          finish(true);
          return [2 /*return*/];
        case 21:
          console.log("[onMessageSendHandler] Skipping pre-flight check, proceeding to upload");
          console.log("[onMessageSendHandler] Reading email metadata");
          console.log("[onMessageSendHandler] Current item info:", {
            itemType: (_l = Office.context.mailbox.item) === null || _l === void 0 ? void 0 : _l.itemType,
            itemClass: (_m = Office.context.mailbox.item) === null || _m === void 0 ? void 0 : _m.itemClass,
            hasSubject: !!((_o = Office.context.mailbox.item) === null || _o === void 0 ? void 0 : _o.subject)
          });
          return [4 /*yield*/, withTimeout(getSubjectRuntime(), T_SUBJECT_MS)];
        case 22:
          subject = _t.sent();
          return [4 /*yield*/, withTimeout(getBodyTextRuntime(), T_BODY_MS)];
        case 23:
          bodyText = _t.sent();
          itemFrom = (_p = Office.context.mailbox.item) === null || _p === void 0 ? void 0 : _p.from;
          fromEmail = String((itemFrom === null || itemFrom === void 0 ? void 0 : itemFrom.emailAddress) || ((_q = Office.context.mailbox.userProfile) === null || _q === void 0 ? void 0 : _q.emailAddress) || "");
          fromName = String((itemFrom === null || itemFrom === void 0 ? void 0 : itemFrom.displayName) || ((_r = Office.context.mailbox.userProfile) === null || _r === void 0 ? void 0 : _r.displayName) || "");
          conversationId = getConversationIdSafe();
          console.log("[onMessageSendHandler] Email metadata", {
            subject: subject,
            fromEmail: fromEmail,
            fromName: fromName,
            bodyLength: bodyText.length,
            hasConversationId: !!conversationId,
            conversationIdPreview: conversationId ? conversationId.substring(0, 30) + "..." : "(none)"
          });
          baseName = safeFileName(subject || "email");
          emailText = "From: ".concat(fromName, " <").concat(fromEmail, ">\r\n") + "To: SingleCase <noreply@singlecase>\r\n" + "Subject: ".concat(subject, "\r\n") + "Date: ".concat(new Date().toUTCString(), "\r\n") + "Message-ID: <".concat(keys[0], "@outlook>\r\n") + "MIME-Version: 1.0\r\n" + "Content-Type: text/plain; charset=UTF-8\r\n" + "Content-Transfer-Encoding: 8bit\r\n" + "\r\n" + "".concat((bodyText || "").trim(), "\r\n");
          emailBase64 = toBase64Utf8(emailText);
          console.log("[onMessageSendHandler] EML built", {
            length: emailBase64.length
          });
          existingDoc = null;
          _t.label = 24;
        case 24:
          _t.trys.push([24, 26,, 27]);
          console.log("[onMessageSendHandler] Checking for existing document with same subject");
          return [4 /*yield*/, withTimeout((0,_services_singlecaseDocuments__WEBPACK_IMPORTED_MODULE_3__.findDocumentBySubject)(intent.caseId, subject), T_FETCH_MS)];
        case 25:
          existingDoc = _t.sent();
          if (existingDoc) {
            console.log("[onMessageSendHandler] Found existing document", {
              docId: existingDoc.id,
              docName: existingDoc.name,
              docSubject: existingDoc.subject
            });
          } else {
            console.log("[onMessageSendHandler] No existing document with this subject found");
          }
          return [3 /*break*/, 27];
        case 26:
          e_5 = _t.sent();
          console.warn("[onMessageSendHandler] Failed to check for existing document:", e_5);
          // Continue with new document creation on error
          existingDoc = null;
          return [3 /*break*/, 27];
        case 27:
          shouldUploadVersion = !!existingDoc;
          console.log("[onMessageSendHandler] Version decision", {
            caseId: intent.caseId,
            subject: subject,
            existingDocId: existingDoc === null || existingDoc === void 0 ? void 0 : existingDoc.id,
            existingDocName: existingDoc === null || existingDoc === void 0 ? void 0 : existingDoc.name,
            shouldUploadVersion: shouldUploadVersion
          });
          if (!(shouldUploadVersion && existingDoc)) return [3 /*break*/, 34];
          // Upload as new version of existing document
          console.log("[onMessageSendHandler] Uploading as version of existing document:", existingDoc.id);
          return [4 /*yield*/, withTimeout((0,_services_singlecaseDocuments__WEBPACK_IMPORTED_MODULE_3__.uploadDocumentVersion)({
            documentId: existingDoc.id,
            fileName: "".concat(baseName, ".eml"),
            mimeType: "message/rfc822",
            dataBase64: emailBase64
          }), T_FETCH_MS)];
        case 28:
          _t.sent();
          console.log("[onMessageSendHandler] Version uploaded successfully");
          // Update filed context with the existing document ID
          return [4 /*yield*/, persistFiledCtx(intent.caseId, existingDoc.id)];
        case 29:
          // Update filed context with the existing document ID
          _t.sent();
          conversationId_1 = getConversationIdSafe();
          if (!conversationId_1) return [3 /*break*/, 31];
          return [4 /*yield*/, (0,_utils_filedCache__WEBPACK_IMPORTED_MODULE_4__.cacheFiledEmail)(conversationId_1, intent.caseId, existingDoc.id, subject)];
        case 30:
          _t.sent();
          console.log("[onMessageSendHandler] Cached filed email (version)", {
            conversationId: conversationId_1.substring(0, 20) + "..."
          });
          return [3 /*break*/, 33];
        case 31:
          // Fallback: Cache by subject when conversationId not available (new compose emails)
          return [4 /*yield*/, (0,_utils_filedCache__WEBPACK_IMPORTED_MODULE_4__.cacheFiledEmailBySubject)(subject, intent.caseId, existingDoc.id)];
        case 32:
          // Fallback: Cache by subject when conversationId not available (new compose emails)
          _t.sent();
          console.log("[onMessageSendHandler] Cached filed email by subject (version)", {
            subject: subject
          });
          _t.label = 33;
        case 33:
          return [3 /*break*/, 40];
        case 34:
          // Upload as new document
          console.log("[onMessageSendHandler] Uploading as new document");
          return [4 /*yield*/, withTimeout((0,_services_singlecaseDocuments__WEBPACK_IMPORTED_MODULE_3__.uploadDocumentToCase)({
            caseId: intent.caseId,
            fileName: "".concat(baseName, ".eml"),
            mimeType: "message/rfc822",
            dataBase64: emailBase64,
            metadata: {
              subject: subject,
              fromEmail: fromEmail,
              fromName: fromName,
              conversationId: conversationId || undefined // Cross-mailbox identifier
            }
          }), T_FETCH_MS)];
        case 35:
          created = _t.sent();
          docs = created === null || created === void 0 ? void 0 : created.documents;
          createdDocId = Array.isArray(docs) && ((_s = docs[0]) === null || _s === void 0 ? void 0 : _s.id) ? String(docs[0].id) : "";
          console.log("[onMessageSendHandler] Created docId", {
            createdDocId: createdDocId,
            rawResponse: created
          });
          if (!createdDocId) return [3 /*break*/, 40];
          return [4 /*yield*/, persistFiledCtx(intent.caseId, createdDocId)];
        case 36:
          _t.sent();
          conversationId_2 = getConversationIdSafe();
          if (!conversationId_2) return [3 /*break*/, 38];
          return [4 /*yield*/, (0,_utils_filedCache__WEBPACK_IMPORTED_MODULE_4__.cacheFiledEmail)(conversationId_2, intent.caseId, createdDocId, subject)];
        case 37:
          _t.sent();
          console.log("[onMessageSendHandler] Cached filed email (new doc)", {
            conversationId: conversationId_2.substring(0, 20) + "..."
          });
          return [3 /*break*/, 40];
        case 38:
          // Fallback: Cache by subject when conversationId not available (new compose emails)
          return [4 /*yield*/, (0,_utils_filedCache__WEBPACK_IMPORTED_MODULE_4__.cacheFiledEmailBySubject)(subject, intent.caseId, createdDocId)];
        case 39:
          // Fallback: Cache by subject when conversationId not available (new compose emails)
          _t.sent();
          console.log("[onMessageSendHandler] Cached filed email by subject (new doc)", {
            subject: subject
          });
          _t.label = 40;
        case 40:
          console.log("[onMessageSendHandler] Upload successful");
          return [4 /*yield*/, showInfo("SingleCase: email uložen při odeslání.")];
        case 41:
          _t.sent();
          finish(true);
          return [3 /*break*/, 47];
        case 42:
          e_6 = _t.sent();
          console.error("[onMessageSendHandler] Error during filing", e_6);
          _t.label = 43;
        case 43:
          _t.trys.push([43, 45,, 46]);
          msg = e_6 instanceof Error ? e_6.message : String(e_6);
          errorHint = "";
          if (msg.includes("timeout")) errorHint = " (timeout)";else if (msg.toLowerCase().includes("workspace")) errorHint = " (není nastaven workspace)";else if (msg.toLowerCase().includes("token")) errorHint = " (přihlaste se znovu)";else if (msg.toLowerCase().includes("network")) errorHint = " (problém se sítí)";
          return [4 /*yield*/, showInfo("SingleCase: nepoda\u0159ilo se ulo\u017Eit".concat(errorHint))];
        case 44:
          _t.sent();
          return [3 /*break*/, 46];
        case 45:
          _a = _t.sent();
          return [3 /*break*/, 46];
        case 46:
          finish(true);
          return [3 /*break*/, 47];
        case 47:
          return [2 /*return*/];
      }
    });
  });
}

/***/ }),

/***/ "./src/services/auth.ts":
/*!******************************!*\
  !*** ./src/services/auth.ts ***!
  \******************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   clearAuth: function() { return /* binding */ clearAuth; },
/* harmony export */   clearAuthIfExpired: function() { return /* binding */ clearAuthIfExpired; },
/* harmony export */   clearAuthIfExpiredRuntime: function() { return /* binding */ clearAuthIfExpiredRuntime; },
/* harmony export */   getAuth: function() { return /* binding */ getAuth; },
/* harmony export */   getAuthRuntime: function() { return /* binding */ getAuthRuntime; },
/* harmony export */   isLoggedIn: function() { return /* binding */ isLoggedIn; },
/* harmony export */   isLoggedInRuntime: function() { return /* binding */ isLoggedInRuntime; },
/* harmony export */   setAuth: function() { return /* binding */ setAuth; }
/* harmony export */ });
/* provided dependency */ var Promise = __webpack_require__(/*! es6-promise */ "./node_modules/es6-promise/dist/es6-promise.js")["Promise"];
/* global OfficeRuntime */
var __awaiter = undefined && undefined.__awaiter || function (thisArg, _arguments, P, generator) {
  function adopt(value) {
    return value instanceof P ? value : new P(function (resolve) {
      resolve(value);
    });
  }
  return new (P || (P = Promise))(function (resolve, reject) {
    function fulfilled(value) {
      try {
        step(generator.next(value));
      } catch (e) {
        reject(e);
      }
    }
    function rejected(value) {
      try {
        step(generator["throw"](value));
      } catch (e) {
        reject(e);
      }
    }
    function step(result) {
      result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected);
    }
    step((generator = generator.apply(thisArg, _arguments || [])).next());
  });
};
var __generator = undefined && undefined.__generator || function (thisArg, body) {
  var _ = {
      label: 0,
      sent: function sent() {
        if (t[0] & 1) throw t[1];
        return t[1];
      },
      trys: [],
      ops: []
    },
    f,
    y,
    t,
    g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
  return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function () {
    return this;
  }), g;
  function verb(n) {
    return function (v) {
      return step([n, v]);
    };
  }
  function step(op) {
    if (f) throw new TypeError("Generator is already executing.");
    while (g && (g = 0, op[0] && (_ = 0)), _) try {
      if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
      if (y = 0, t) op = [op[0] & 2, t.value];
      switch (op[0]) {
        case 0:
        case 1:
          t = op;
          break;
        case 4:
          _.label++;
          return {
            value: op[1],
            done: false
          };
        case 5:
          _.label++;
          y = op[1];
          op = [0];
          continue;
        case 7:
          op = _.ops.pop();
          _.trys.pop();
          continue;
        default:
          if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) {
            _ = 0;
            continue;
          }
          if (op[0] === 3 && (!t || op[1] > t[0] && op[1] < t[3])) {
            _.label = op[1];
            break;
          }
          if (op[0] === 6 && _.label < t[1]) {
            _.label = t[1];
            t = op;
            break;
          }
          if (t && _.label < t[2]) {
            _.label = t[2];
            _.ops.push(op);
            break;
          }
          if (t[2]) _.ops.pop();
          _.trys.pop();
          continue;
      }
      op = body.call(thisArg, _);
    } catch (e) {
      op = [6, e];
      y = 0;
    } finally {
      f = t = 0;
    }
    if (op[0] & 5) throw op[1];
    return {
      value: op[0] ? op[1] : void 0,
      done: true
    };
  }
};
var TOKEN_KEY = "singlecase_token";
var USER_KEY = "singlecase_user_email";
var ISSUED_AT_KEY = "singlecase_auth_issued_at";
// Mirror keys into OfficeRuntime.storage so Commands runtime can read them
var RT_TOKEN_KEY = "sc_token";
var RT_USER_KEY = "sc_user_email";
var RT_ISSUED_AT_KEY = "sc_auth_issued_at";
// Typical session TTL: 8 hours. Adjust as you like.
var SESSION_TTL_MS = 8 * 60 * 60 * 1000;
function normalizeEmail(email) {
  var v = (email || "").trim().toLowerCase();
  return v.length > 0 ? v : "unknown@singlecase.local";
}
function rtGet(key) {
  return __awaiter(this, void 0, void 0, function () {
    var v, e_1, v;
    var _a;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          if (!(typeof OfficeRuntime !== 'undefined' && (OfficeRuntime === null || OfficeRuntime === void 0 ? void 0 : OfficeRuntime.storage))) return [3 /*break*/, 4];
          _b.label = 1;
        case 1:
          _b.trys.push([1, 3,, 4]);
          return [4 /*yield*/, OfficeRuntime.storage.getItem(key)];
        case 2:
          v = _b.sent();
          if (typeof v === "string") return [2 /*return*/, v];
          return [3 /*break*/, 4];
        case 3:
          e_1 = _b.sent();
          console.warn("[rtGet] OfficeRuntime.storage.getItem failed:", e_1);
          return [3 /*break*/, 4];
        case 4:
          // Fallback to Office.context.roamingSettings (Outlook-specific, works cross-context)
          if ((_a = Office === null || Office === void 0 ? void 0 : Office.context) === null || _a === void 0 ? void 0 : _a.roamingSettings) {
            try {
              v = Office.context.roamingSettings.get(key);
              if (typeof v === "string") return [2 /*return*/, v];
            } catch (e) {
              console.warn("[rtGet] roamingSettings.get failed:", e);
            }
          }
          return [2 /*return*/, null];
      }
    });
  });
}
function rtSet(key, value) {
  return __awaiter(this, void 0, void 0, function () {
    var e_2, e_3;
    var _a;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          if (!(typeof OfficeRuntime !== 'undefined' && (OfficeRuntime === null || OfficeRuntime === void 0 ? void 0 : OfficeRuntime.storage))) return [3 /*break*/, 4];
          _b.label = 1;
        case 1:
          _b.trys.push([1, 3,, 4]);
          return [4 /*yield*/, OfficeRuntime.storage.setItem(key, value)];
        case 2:
          _b.sent();
          console.log("[rtSet] Saved to OfficeRuntime.storage:", key);
          return [2 /*return*/];
        // Success
        case 3:
          e_2 = _b.sent();
          console.warn("[rtSet] OfficeRuntime.storage.setItem failed:", e_2);
          return [3 /*break*/, 4];
        case 4:
          if (!((_a = Office === null || Office === void 0 ? void 0 : Office.context) === null || _a === void 0 ? void 0 : _a.roamingSettings)) return [3 /*break*/, 8];
          _b.label = 5;
        case 5:
          _b.trys.push([5, 7,, 8]);
          Office.context.roamingSettings.set(key, value);
          return [4 /*yield*/, new Promise(function (resolve, reject) {
            Office.context.roamingSettings.saveAsync(function (result) {
              var _a;
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("[rtSet] Saved to roamingSettings:", key);
                resolve();
              } else {
                console.error("[rtSet] roamingSettings.saveAsync failed:", result.error);
                reject(new Error(((_a = result.error) === null || _a === void 0 ? void 0 : _a.message) || "saveAsync failed"));
              }
            });
          })];
        case 6:
          _b.sent();
          return [3 /*break*/, 8];
        case 7:
          e_3 = _b.sent();
          console.error("[rtSet] roamingSettings failed:", e_3);
          return [3 /*break*/, 8];
        case 8:
          return [2 /*return*/];
      }
    });
  });
}
function rtRemove(key) {
  return __awaiter(this, void 0, void 0, function () {
    var _a;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          _b.trys.push([0, 2,, 3]);
          return [4 /*yield*/, OfficeRuntime.storage.removeItem(key)];
        case 1:
          _b.sent();
          return [3 /*break*/, 3];
        case 2:
          _a = _b.sent();
          return [3 /*break*/, 3];
        case 3:
          return [2 /*return*/];
      }
    });
  });
}
function getAuth() {
  var token = sessionStorage.getItem(TOKEN_KEY);
  var emailRaw = sessionStorage.getItem(USER_KEY);
  var issuedAtStr = sessionStorage.getItem(ISSUED_AT_KEY);
  return {
    token: token,
    email: normalizeEmail(emailRaw),
    issuedAt: issuedAtStr ? Number(issuedAtStr) : 0
  };
}
// Async version for runtimes that cannot access sessionStorage (eg Commands)
function getAuthRuntime() {
  return __awaiter(this, void 0, void 0, function () {
    var _a, token, emailRaw, issuedAtStr;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          return [4 /*yield*/, Promise.all([rtGet(RT_TOKEN_KEY), rtGet(RT_USER_KEY), rtGet(RT_ISSUED_AT_KEY)])];
        case 1:
          _a = _b.sent(), token = _a[0], emailRaw = _a[1], issuedAtStr = _a[2];
          return [2 /*return*/, {
            token: token,
            email: normalizeEmail(emailRaw),
            issuedAt: issuedAtStr ? Number(issuedAtStr) : 0
          }];
      }
    });
  });
}
// Make this async so you can await the mirror write when needed.
function setAuth(token, email) {
  return __awaiter(this, void 0, void 0, function () {
    var emailNorm, issuedAt;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          emailNorm = normalizeEmail(email);
          issuedAt = Date.now();
          sessionStorage.setItem(TOKEN_KEY, token);
          sessionStorage.setItem(USER_KEY, emailNorm);
          sessionStorage.setItem(ISSUED_AT_KEY, String(issuedAt));
          // Mirror for command runtime
          return [4 /*yield*/, Promise.all([rtSet(RT_TOKEN_KEY, token), rtSet(RT_USER_KEY, emailNorm), rtSet(RT_ISSUED_AT_KEY, String(issuedAt))])];
        case 1:
          // Mirror for command runtime
          _a.sent();
          return [2 /*return*/];
      }
    });
  });
}
function clearAuthIfExpired() {
  var _a = getAuth(),
    token = _a.token,
    issuedAt = _a.issuedAt;
  if (!token) return;
  var ageMs = Date.now() - (issuedAt || 0);
  if (!issuedAt || ageMs > SESSION_TTL_MS) {
    void clearAuth();
  }
}
function clearAuthIfExpiredRuntime() {
  return __awaiter(this, void 0, void 0, function () {
    var _a, token, issuedAt, ageMs;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          return [4 /*yield*/, getAuthRuntime()];
        case 1:
          _a = _b.sent(), token = _a.token, issuedAt = _a.issuedAt;
          if (!token) return [2 /*return*/];
          ageMs = Date.now() - (issuedAt || 0);
          if (!(!issuedAt || ageMs > SESSION_TTL_MS)) return [3 /*break*/, 3];
          return [4 /*yield*/, clearAuth()];
        case 2:
          _b.sent();
          _b.label = 3;
        case 3:
          return [2 /*return*/];
      }
    });
  });
}
function clearAuth() {
  return __awaiter(this, void 0, void 0, function () {
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          sessionStorage.removeItem(TOKEN_KEY);
          sessionStorage.removeItem(USER_KEY);
          sessionStorage.removeItem(ISSUED_AT_KEY);
          return [4 /*yield*/, Promise.all([rtRemove(RT_TOKEN_KEY), rtRemove(RT_USER_KEY), rtRemove(RT_ISSUED_AT_KEY)])];
        case 1:
          _a.sent();
          return [2 /*return*/];
      }
    });
  });
}
function isLoggedIn() {
  var _a = getAuth(),
    token = _a.token,
    issuedAt = _a.issuedAt;
  if (!token) return false;
  var ageMs = Date.now() - (issuedAt || 0);
  return Boolean(issuedAt && ageMs <= SESSION_TTL_MS);
}
function isLoggedInRuntime() {
  return __awaiter(this, void 0, void 0, function () {
    var _a, token, issuedAt, ageMs;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          return [4 /*yield*/, getAuthRuntime()];
        case 1:
          _a = _b.sent(), token = _a.token, issuedAt = _a.issuedAt;
          if (!token) return [2 /*return*/, false];
          ageMs = Date.now() - (issuedAt || 0);
          return [2 /*return*/, Boolean(issuedAt && ageMs <= SESSION_TTL_MS)];
      }
    });
  });
}

/***/ }),

/***/ "./src/services/singlecaseDocuments.ts":
/*!*********************************************!*\
  !*** ./src/services/singlecaseDocuments.ts ***!
  \*********************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   checkFiledStatusByConversationAndSubject: function() { return /* binding */ checkFiledStatusByConversationAndSubject; },
/* harmony export */   createDirectory: function() { return /* binding */ createDirectory; },
/* harmony export */   ensureOutlookAddinFolder: function() { return /* binding */ ensureOutlookAddinFolder; },
/* harmony export */   findDocumentBySubject: function() { return /* binding */ findDocumentBySubject; },
/* harmony export */   getCaseRootDirectoryId: function() { return /* binding */ getCaseRootDirectoryId; },
/* harmony export */   getDocumentMeta: function() { return /* binding */ getDocumentMeta; },
/* harmony export */   listDirectory: function() { return /* binding */ listDirectory; },
/* harmony export */   normalizeSubject: function() { return /* binding */ normalizeSubject; },
/* harmony export */   renameDocument: function() { return /* binding */ renameDocument; },
/* harmony export */   uploadDocumentToCase: function() { return /* binding */ uploadDocumentToCase; },
/* harmony export */   uploadDocumentVersion: function() { return /* binding */ uploadDocumentVersion; }
/* harmony export */ });
/* harmony import */ var _auth__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./auth */ "./src/services/auth.ts");
/* harmony import */ var _utils_storage__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../utils/storage */ "./src/utils/storage.ts");
/* harmony import */ var _utils_constants__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utils/constants */ "./src/utils/constants.ts");
/* provided dependency */ var Promise = __webpack_require__(/*! es6-promise */ "./node_modules/es6-promise/dist/es6-promise.js")["Promise"];
var __assign = undefined && undefined.__assign || function () {
  __assign = Object.assign || function (t) {
    for (var s, i = 1, n = arguments.length; i < n; i++) {
      s = arguments[i];
      for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
    }
    return t;
  };
  return __assign.apply(this, arguments);
};
var __awaiter = undefined && undefined.__awaiter || function (thisArg, _arguments, P, generator) {
  function adopt(value) {
    return value instanceof P ? value : new P(function (resolve) {
      resolve(value);
    });
  }
  return new (P || (P = Promise))(function (resolve, reject) {
    function fulfilled(value) {
      try {
        step(generator.next(value));
      } catch (e) {
        reject(e);
      }
    }
    function rejected(value) {
      try {
        step(generator["throw"](value));
      } catch (e) {
        reject(e);
      }
    }
    function step(result) {
      result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected);
    }
    step((generator = generator.apply(thisArg, _arguments || [])).next());
  });
};
var __generator = undefined && undefined.__generator || function (thisArg, body) {
  var _ = {
      label: 0,
      sent: function sent() {
        if (t[0] & 1) throw t[1];
        return t[1];
      },
      trys: [],
      ops: []
    },
    f,
    y,
    t,
    g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
  return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function () {
    return this;
  }), g;
  function verb(n) {
    return function (v) {
      return step([n, v]);
    };
  }
  function step(op) {
    if (f) throw new TypeError("Generator is already executing.");
    while (g && (g = 0, op[0] && (_ = 0)), _) try {
      if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
      if (y = 0, t) op = [op[0] & 2, t.value];
      switch (op[0]) {
        case 0:
        case 1:
          t = op;
          break;
        case 4:
          _.label++;
          return {
            value: op[1],
            done: false
          };
        case 5:
          _.label++;
          y = op[1];
          op = [0];
          continue;
        case 7:
          op = _.ops.pop();
          _.trys.pop();
          continue;
        default:
          if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) {
            _ = 0;
            continue;
          }
          if (op[0] === 3 && (!t || op[1] > t[0] && op[1] < t[3])) {
            _.label = op[1];
            break;
          }
          if (op[0] === 6 && _.label < t[1]) {
            _.label = t[1];
            t = op;
            break;
          }
          if (t && _.label < t[2]) {
            _.label = t[2];
            _.ops.push(op);
            break;
          }
          if (t[2]) _.ops.pop();
          _.trys.pop();
          continue;
      }
      op = body.call(thisArg, _);
    } catch (e) {
      op = [6, e];
      y = 0;
    } finally {
      f = t = 0;
    }
    if (op[0] & 5) throw op[1];
    return {
      value: op[0] ? op[1] : void 0,
      done: true
    };
  }
};



function normalizeHost(host) {
  var v = (host || "").trim().toLowerCase();
  if (!v) return "";
  return v.replace(/^https?:\/\//i, "").split("/")[0];
}
function resolveApiBaseUrl() {
  return __awaiter(this, void 0, void 0, function () {
    var storedHostRaw, host, baseUrl;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          console.log("[resolveApiBaseUrl] Reading workspaceHost from storage key:", _utils_constants__WEBPACK_IMPORTED_MODULE_2__.STORAGE_KEYS.workspaceHost);
          return [4 /*yield*/, (0,_utils_storage__WEBPACK_IMPORTED_MODULE_1__.getStored)(_utils_constants__WEBPACK_IMPORTED_MODULE_2__.STORAGE_KEYS.workspaceHost)];
        case 1:
          storedHostRaw = _a.sent();
          console.log("[resolveApiBaseUrl] Raw stored host:", storedHostRaw);
          host = normalizeHost(storedHostRaw || "");
          console.log("[resolveApiBaseUrl] Normalized host:", host);
          if (!host) {
            console.error("[resolveApiBaseUrl] Workspace host is missing");
            throw new Error("Workspace host is missing.");
          }
          baseUrl = "/singlecase/".concat(encodeURIComponent(host), "/publicapi/v1");
          console.log("[resolveApiBaseUrl] Resolved base URL:", baseUrl);
          return [2 /*return*/, baseUrl];
      }
    });
  });
}
function getToken() {
  return __awaiter(this, void 0, void 0, function () {
    var auth, rt;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          auth = (0,_auth__WEBPACK_IMPORTED_MODULE_0__.getAuth)();
          if (auth === null || auth === void 0 ? void 0 : auth.token) {
            console.log("[getToken] Using sessionStorage token");
            return [2 /*return*/, auth.token];
          }
          console.log("[getToken] sessionStorage token not available, trying OfficeRuntime.storage");
          return [4 /*yield*/, (0,_auth__WEBPACK_IMPORTED_MODULE_0__.getAuthRuntime)()];
        case 1:
          rt = _a.sent();
          if (rt === null || rt === void 0 ? void 0 : rt.token) {
            console.log("[getToken] Using OfficeRuntime.storage token");
            return [2 /*return*/, rt.token];
          }
          console.error("[getToken] No token found in either sessionStorage or OfficeRuntime.storage");
          throw new Error("Missing auth token.");
      }
    });
  });
}
function expectJson(res, errorPrefix) {
  return __awaiter(this, void 0, void 0, function () {
    var text, contentType;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          return [4 /*yield*/, res.text().catch(function () {
            return "";
          })];
        case 1:
          text = _a.sent();
          if (!res.ok) {
            if (res.status === 423) {
              throw new Error("Dokument je momentálně uzamčen. Někdo jej právě upravuje. Počkejte prosím, než se dokument odemkne, a zkuste to znovu.");
            }
            throw new Error("".concat(errorPrefix, " (").concat(res.status, "): ").concat(text || res.statusText));
          }
          contentType = res.headers.get("content-type") || "";
          if (!contentType.includes("application/json")) {
            throw new Error("".concat(errorPrefix, ": expected JSON but got ").concat(contentType || "no content-type", "."));
          }
          return [2 /*return*/, JSON.parse(text)];
      }
    });
  });
}
function getDocumentMeta(documentId) {
  return __awaiter(this, void 0, void 0, function () {
    var token, base, url, res, json;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          return [4 /*yield*/, getToken()];
        case 1:
          token = _a.sent();
          return [4 /*yield*/, resolveApiBaseUrl()];
        case 2:
          base = _a.sent();
          url = "".concat(base, "/documents/").concat(encodeURIComponent(String(documentId)));
          return [4 /*yield*/, fetch(url, {
            method: "GET",
            headers: {
              Authentication: token,
              "Content-Type": "application/json",
              "Accept-Encoding": "identity"
            }
          })];
        case 3:
          res = _a.sent();
          if (res.status === 404) return [2 /*return*/, null];
          return [4 /*yield*/, expectJson(res, "Get document failed")];
        case 4:
          json = _a.sent();
          return [2 /*return*/, {
            id: String(json.id || documentId),
            name: String(json.name || ""),
            case_id: String(json.case_id || "")
          }];
      }
    });
  });
}
function uploadDocumentVersion(params) {
  return __awaiter(this, void 0, void 0, function () {
    var documentId, fileName, mimeType, dataBase64, directoryId, token, base, id, bodyData, body, candidates, lastErr, _i, candidates_1, c, res, json;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          documentId = params.documentId, fileName = params.fileName, mimeType = params.mimeType, dataBase64 = params.dataBase64, directoryId = params.directoryId;
          return [4 /*yield*/, getToken()];
        case 1:
          token = _a.sent();
          return [4 /*yield*/, resolveApiBaseUrl()];
        case 2:
          base = _a.sent();
          id = encodeURIComponent(String(documentId));
          bodyData = {
            name: fileName,
            mime_type: mimeType,
            data_base64: dataBase64
          };
          // Add dir_id if provided (though versions typically inherit parent directory)
          if (directoryId) {
            bodyData.dir_id = directoryId;
          }
          body = JSON.stringify(bodyData);
          candidates = [{
            url: "".concat(base, "/documents/").concat(id, "/version"),
            method: "POST"
          }, {
            url: "".concat(base, "/documents/").concat(id, "/versions"),
            method: "POST"
          }, {
            url: "".concat(base, "/documents/").concat(id, "/versions"),
            method: "PUT"
          }, {
            url: "".concat(base, "/documents/").concat(id, "/version"),
            method: "PUT"
          }, {
            url: "".concat(base, "/documents/").concat(id, "/versions"),
            method: "PATCH"
          }, {
            url: "".concat(base, "/documents/").concat(id, "/version"),
            method: "PATCH"
          }];
          lastErr = null;
          _i = 0, candidates_1 = candidates;
          _a.label = 3;
        case 3:
          if (!(_i < candidates_1.length)) return [3 /*break*/, 7];
          c = candidates_1[_i];
          return [4 /*yield*/, fetch(c.url, {
            method: c.method,
            headers: {
              "Content-Type": "application/json",
              Authentication: token,
              "Accept-Encoding": "identity"
            },
            body: body
          })];
        case 4:
          res = _a.sent();
          if (res.status === 404 || res.status === 405) {
            lastErr = new Error("Endpoint not available: ".concat(c.method, " ").concat(c.url, " (").concat(res.status, ")"));
            return [3 /*break*/, 6];
          }
          return [4 /*yield*/, expectJson(res, "Upload version failed")];
        case 5:
          json = _a.sent();
          return [2 /*return*/, json];
        case 6:
          _i++;
          return [3 /*break*/, 3];
        case 7:
          throw lastErr instanceof Error ? lastErr : new Error("Upload version failed: no supported endpoint found");
      }
    });
  });
}
function uploadDocumentToCase(params) {
  return __awaiter(this, void 0, void 0, function () {
    var caseId, fileName, mimeType, dataBase64, directoryId, metadata, token, e_1, base, e_2, url, payload, res, e_3, text, snippet, json;
    var _a;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          caseId = params.caseId, fileName = params.fileName, mimeType = params.mimeType, dataBase64 = params.dataBase64, directoryId = params.directoryId, metadata = params.metadata;
          console.log("[uploadDocumentToCase] Starting upload", {
            caseId: caseId,
            fileName: fileName,
            mimeType: mimeType,
            dataLength: dataBase64.length
          });
          _b.label = 1;
        case 1:
          _b.trys.push([1, 3,, 4]);
          return [4 /*yield*/, getToken()];
        case 2:
          token = _b.sent();
          console.log("[uploadDocumentToCase] Token retrieved", {
            hasToken: !!token,
            tokenPrefix: token.slice(0, 10)
          });
          return [3 /*break*/, 4];
        case 3:
          e_1 = _b.sent();
          console.error("[uploadDocumentToCase] Failed to get token:", e_1);
          throw e_1;
        case 4:
          _b.trys.push([4, 6,, 7]);
          return [4 /*yield*/, resolveApiBaseUrl()];
        case 5:
          base = _b.sent();
          console.log("[uploadDocumentToCase] Base URL resolved:", base);
          return [3 /*break*/, 7];
        case 6:
          e_2 = _b.sent();
          console.error("[uploadDocumentToCase] Failed to resolve base URL:", e_2);
          throw e_2;
        case 7:
          url = "".concat(base, "/documents");
          console.log("[uploadDocumentToCase] Full URL:", url);
          payload = {
            case_id: caseId,
            documents: [__assign(__assign({
              name: fileName,
              mime_type: mimeType,
              data_base64: dataBase64
            }, directoryId ? {
              dir_id: directoryId
            } : {}), metadata ? {
              metadata: metadata
            } : {})]
          };
          console.log("[uploadDocumentToCase] Payload structure:", {
            case_id: payload.case_id,
            documentCount: payload.documents.length,
            firstDoc: {
              name: payload.documents[0].name,
              mime_type: payload.documents[0].mime_type,
              data_base64_length: payload.documents[0].data_base64.length
            }
          });
          _b.label = 8;
        case 8:
          _b.trys.push([8, 10,, 11]);
          return [4 /*yield*/, fetch(url, {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              Authentication: token,
              "Accept-Encoding": "identity"
            },
            body: JSON.stringify(payload)
          })];
        case 9:
          res = _b.sent();
          console.log("[uploadDocumentToCase] Fetch completed", {
            status: res.status,
            statusText: res.statusText,
            ok: res.ok
          });
          return [3 /*break*/, 11];
        case 10:
          e_3 = _b.sent();
          console.error("[uploadDocumentToCase] Fetch failed:", e_3);
          throw new Error("Network request failed: ".concat(e_3 instanceof Error ? e_3.message : String(e_3)));
        case 11:
          if (!!res.ok) return [3 /*break*/, 13];
          return [4 /*yield*/, res.text().catch(function () {
            return "";
          })];
        case 12:
          text = _b.sent();
          snippet = text.slice(0, 300);
          console.error("[uploadDocumentToCase] Upload failed", {
            status: res.status,
            statusText: res.statusText,
            url: url,
            responseSnippet: snippet
          });
          _b.label = 13;
        case 13:
          return [4 /*yield*/, expectJson(res, "Upload failed")];
        case 14:
          json = _b.sent();
          console.log("[uploadDocumentToCase] Upload successful", {
            documentIds: (_a = json.documents) === null || _a === void 0 ? void 0 : _a.map(function (d) {
              return d.id;
            })
          });
          return [2 /*return*/, json];
      }
    });
  });
}
function tryRename(url, method, token, name) {
  return __awaiter(this, void 0, void 0, function () {
    var res, contentType, text;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          return [4 /*yield*/, fetch(url, {
            method: method,
            headers: {
              "Content-Type": "application/json",
              Authentication: token,
              "Accept-Encoding": "identity"
            },
            body: JSON.stringify({
              name: name
            })
          })];
        case 1:
          res = _a.sent();
          if (res.ok) {
            contentType = res.headers.get("content-type") || "";
            if (contentType.includes("application/json")) return [2 /*return*/, res.json()];
            return [2 /*return*/, {}];
          }
          if (res.status === 404 || res.status === 405) return [2 /*return*/, null];
          return [4 /*yield*/, res.text().catch(function () {
            return "";
          })];
        case 2:
          text = _a.sent();
          throw new Error("Rename document failed (".concat(res.status, "): ").concat(text || res.statusText));
      }
    });
  });
}
function renameDocument(params) {
  return __awaiter(this, void 0, void 0, function () {
    var base, id, candidates, _i, candidates_2, c, ok;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          return [4 /*yield*/, resolveApiBaseUrl()];
        case 1:
          base = _a.sent();
          id = encodeURIComponent(String(params.documentId));
          candidates = [{
            url: "".concat(base, "/documents/").concat(id),
            method: "PUT"
          }, {
            url: "".concat(base, "/documents/").concat(id),
            method: "PATCH"
          }, {
            url: "".concat(base, "/documents/").concat(id, "/rename"),
            method: "POST"
          }, {
            url: "".concat(base, "/documents/").concat(id, "/name"),
            method: "PUT"
          }];
          _i = 0, candidates_2 = candidates;
          _a.label = 2;
        case 2:
          if (!(_i < candidates_2.length)) return [3 /*break*/, 5];
          c = candidates_2[_i];
          return [4 /*yield*/, tryRename(c.url, c.method, params.token, params.name)];
        case 3:
          ok = _a.sent();
          if (ok) return [2 /*return*/, ok];
          _a.label = 4;
        case 4:
          _i++;
          return [3 /*break*/, 2];
        case 5:
          throw new Error("Rename document failed: API endpoint not found or not allowed");
      }
    });
  });
}
// ============================================================================
// Folder/Directory Management for "Outlook add-in" folder
// ============================================================================
var OUTLOOK_FOLDER_NAME = "Outlook add-in";
/**
 * Get cached folder ID for a case, if available
 */
function getCachedFolderId(caseId) {
  return __awaiter(this, void 0, void 0, function () {
    var raw, cache, folderId, _a;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          _b.trys.push([0, 2,, 3]);
          return [4 /*yield*/, (0,_utils_storage__WEBPACK_IMPORTED_MODULE_1__.getStored)(_utils_constants__WEBPACK_IMPORTED_MODULE_2__.STORAGE_KEYS.outlookFolderCache)];
        case 1:
          raw = _b.sent();
          if (!raw) return [2 /*return*/, null];
          cache = JSON.parse(String(raw));
          folderId = cache[String(caseId)];
          return [2 /*return*/, folderId ? String(folderId) : null];
        case 2:
          _a = _b.sent();
          return [2 /*return*/, null];
        case 3:
          return [2 /*return*/];
      }
    });
  });
}
/**
 * Cache folder ID for a case
 */
function cacheFolderId(caseId, folderId) {
  return __awaiter(this, void 0, void 0, function () {
    var raw, cache, e_4;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          _a.trys.push([0, 3,, 4]);
          return [4 /*yield*/, (0,_utils_storage__WEBPACK_IMPORTED_MODULE_1__.getStored)(_utils_constants__WEBPACK_IMPORTED_MODULE_2__.STORAGE_KEYS.outlookFolderCache)];
        case 1:
          raw = _a.sent();
          cache = raw ? JSON.parse(String(raw)) : {};
          cache[String(caseId)] = String(folderId);
          return [4 /*yield*/, (0,_utils_storage__WEBPACK_IMPORTED_MODULE_1__.setStored)(_utils_constants__WEBPACK_IMPORTED_MODULE_2__.STORAGE_KEYS.outlookFolderCache, JSON.stringify(cache))];
        case 2:
          _a.sent();
          return [3 /*break*/, 4];
        case 3:
          e_4 = _a.sent();
          console.warn("[cacheFolderId] Failed to cache folder ID:", e_4);
          return [3 /*break*/, 4];
        case 4:
          return [2 /*return*/];
      }
    });
  });
}
/**
 * Get the root directory ID for a case
 * Returns null if case has no documents directory yet
 */
function getCaseRootDirectoryId(caseId) {
  return __awaiter(this, void 0, void 0, function () {
    var token, base, url, res, json, rootDirId;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          return [4 /*yield*/, getToken()];
        case 1:
          token = _a.sent();
          return [4 /*yield*/, resolveApiBaseUrl()];
        case 2:
          base = _a.sent();
          url = "".concat(base, "/cases/").concat(encodeURIComponent(caseId));
          return [4 /*yield*/, fetch(url, {
            method: "GET",
            headers: {
              Authentication: token,
              "Content-Type": "application/json",
              "Accept-Encoding": "identity"
            }
          })];
        case 3:
          res = _a.sent();
          if (!res.ok) {
            console.warn("[getCaseRootDirectoryId] Failed to get case:", res.status);
            return [2 /*return*/, null];
          }
          return [4 /*yield*/, res.json()];
        case 4:
          json = _a.sent();
          rootDirId = (json === null || json === void 0 ? void 0 : json.root_directory_id) || (json === null || json === void 0 ? void 0 : json.documents_directory_id);
          return [2 /*return*/, rootDirId ? String(rootDirId) : null];
      }
    });
  });
}
/**
 * List contents of a directory
 */
function listDirectory(directoryId) {
  return __awaiter(this, void 0, void 0, function () {
    var token, base, url, res, json, items, rawItems, _i, rawItems_1, item;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          return [4 /*yield*/, getToken()];
        case 1:
          token = _a.sent();
          return [4 /*yield*/, resolveApiBaseUrl()];
        case 2:
          base = _a.sent();
          url = "".concat(base, "/directories/").concat(encodeURIComponent(directoryId));
          return [4 /*yield*/, fetch(url, {
            method: "GET",
            headers: {
              Authentication: token,
              "Content-Type": "application/json",
              "Accept-Encoding": "identity"
            }
          })];
        case 3:
          res = _a.sent();
          if (!res.ok) {
            throw new Error("List directory failed (".concat(res.status, "): ").concat(res.statusText));
          }
          return [4 /*yield*/, res.json()];
        case 4:
          json = _a.sent();
          items = [];
          rawItems = (json === null || json === void 0 ? void 0 : json.items) || (json === null || json === void 0 ? void 0 : json.children) || [];
          for (_i = 0, rawItems_1 = rawItems; _i < rawItems_1.length; _i++) {
            item = rawItems_1[_i];
            items.push({
              id: String(item.id || item._id),
              name: String(item.name || ""),
              type: item.type === "directory" || item.is_directory ? "directory" : "file",
              parent_id: item.parent_id
            });
          }
          return [2 /*return*/, {
            items: items,
            parent_id: json === null || json === void 0 ? void 0 : json.id
          }];
      }
    });
  });
}
/**
 * Create a new directory
 */
function createDirectory(parentId, name) {
  return __awaiter(this, void 0, void 0, function () {
    var token, base, payload, candidates, lastError, _i, candidates_3, candidate, res, text, json, e_5;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          return [4 /*yield*/, getToken()];
        case 1:
          token = _a.sent();
          return [4 /*yield*/, resolveApiBaseUrl()];
        case 2:
          base = _a.sent();
          payload = {
            name: name,
            parent_id: parentId
          };
          candidates = [{
            url: "".concat(base, "/directories"),
            method: "POST"
          }, {
            url: "".concat(base, "/directories/").concat(encodeURIComponent(parentId), "/subdirectories"),
            method: "POST"
          }, {
            url: "".concat(base, "/folders"),
            method: "POST"
          }];
          lastError = null;
          _i = 0, candidates_3 = candidates;
          _a.label = 3;
        case 3:
          if (!(_i < candidates_3.length)) return [3 /*break*/, 11];
          candidate = candidates_3[_i];
          _a.label = 4;
        case 4:
          _a.trys.push([4, 9,, 10]);
          return [4 /*yield*/, fetch(candidate.url, {
            method: candidate.method,
            headers: {
              Authentication: token,
              "Content-Type": "application/json",
              "Accept-Encoding": "identity"
            },
            body: JSON.stringify(payload)
          })];
        case 5:
          res = _a.sent();
          if (res.status === 404 || res.status === 405) {
            return [3 /*break*/, 10]; // Try next endpoint
          }
          if (!!res.ok) return [3 /*break*/, 7];
          return [4 /*yield*/, res.text().catch(function () {
            return "";
          })];
        case 6:
          text = _a.sent();
          lastError = new Error("Create directory failed (".concat(res.status, "): ").concat(text || res.statusText));
          return [3 /*break*/, 10];
        case 7:
          return [4 /*yield*/, res.json()];
        case 8:
          json = _a.sent();
          return [2 /*return*/, {
            id: String(json.id || json._id),
            name: String(json.name || name),
            parent_id: json.parent_id
          }];
        case 9:
          e_5 = _a.sent();
          lastError = e_5 instanceof Error ? e_5 : new Error(String(e_5));
          return [3 /*break*/, 10];
        case 10:
          _i++;
          return [3 /*break*/, 3];
        case 11:
          throw lastError || new Error("Create directory failed: no supported endpoint found");
      }
    });
  });
}
/**
 * Ensure the "Outlook add-in" folder exists in the case
 * Returns the folder's directory ID
 *
 * This function is idempotent and handles concurrent calls safely
 */
function ensureOutlookAddinFolder(caseId) {
  return __awaiter(this, void 0, void 0, function () {
    var cachedId, rootDirId, listing, existing, created, e_6;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          console.log("[ensureOutlookAddinFolder] Starting for case:", caseId);
          return [4 /*yield*/, getCachedFolderId(caseId)];
        case 1:
          cachedId = _a.sent();
          if (cachedId) {
            console.log("[ensureOutlookAddinFolder] Using cached folder ID:", cachedId);
            return [2 /*return*/, cachedId];
          }
          return [4 /*yield*/, getCaseRootDirectoryId(caseId)];
        case 2:
          rootDirId = _a.sent();
          if (!rootDirId) {
            console.warn("[ensureOutlookAddinFolder] Case has no root directory");
            return [2 /*return*/, null];
          }
          console.log("[ensureOutlookAddinFolder] Root directory ID:", rootDirId);
          _a.label = 3;
        case 3:
          _a.trys.push([3, 9,, 10]);
          return [4 /*yield*/, listDirectory(rootDirId)];
        case 4:
          listing = _a.sent();
          console.log("[ensureOutlookAddinFolder] Found", listing.items.length, "items in root");
          existing = listing.items.find(function (item) {
            return item.type === "directory" && item.name === OUTLOOK_FOLDER_NAME;
          });
          if (!existing) return [3 /*break*/, 6];
          console.log("[ensureOutlookAddinFolder] Folder already exists:", existing.id);
          return [4 /*yield*/, cacheFolderId(caseId, String(existing.id))];
        case 5:
          _a.sent();
          return [2 /*return*/, String(existing.id)];
        case 6:
          // 5. Create the folder
          console.log("[ensureOutlookAddinFolder] Creating folder:", OUTLOOK_FOLDER_NAME);
          return [4 /*yield*/, createDirectory(rootDirId, OUTLOOK_FOLDER_NAME)];
        case 7:
          created = _a.sent();
          console.log("[ensureOutlookAddinFolder] Folder created:", created.id);
          return [4 /*yield*/, cacheFolderId(caseId, String(created.id))];
        case 8:
          _a.sent();
          return [2 /*return*/, String(created.id)];
        case 9:
          e_6 = _a.sent();
          console.error("[ensureOutlookAddinFolder] Failed:", e_6);
          // Return null - uploads will go to root if folder creation fails
          return [2 /*return*/, null];
        case 10:
          return [2 /*return*/];
      }
    });
  });
}
// ============================================================================
// Subject-Based Document Matching (for email versioning)
// ============================================================================
/**
 * Normalize email subject for matching
 * - Trim whitespace
 * - Lowercase
 * - Collapse multiple spaces
 * - Optionally strip Re:/Fw:/Fwd: prefixes
 */
function normalizeSubject(subject, stripPrefixes) {
  if (stripPrefixes === void 0) {
    stripPrefixes = true;
  }
  if (!subject) return "";
  var normalized = subject.trim().toLowerCase();
  // Collapse multiple spaces to single space
  normalized = normalized.replace(/\s+/g, " ");
  // Optionally strip common reply/forward prefixes
  if (stripPrefixes) {
    // Remove Re:, RE:, Fw:, FW:, Fwd:, FWD: (with optional spaces and colons)
    // Handle multiple nested prefixes like "Re: Fw: Re: Subject"
    var prevLength = void 0;
    do {
      prevLength = normalized.length;
      normalized = normalized.replace(/^(re|fw|fwd):\s*/i, "");
    } while (normalized.length !== prevLength && normalized.length > 0);
  }
  return normalized.trim();
}
/**
 * Search for existing email documents in a case by matching subject
 * Returns the document ID if found, null otherwise
 *
 * This function:
 * 1. Lists all documents in the case
 * 2. Filters for .eml files
 * 3. Compares normalized subjects
 * 4. Returns the first match (or null)
 */
function findDocumentBySubject(caseId, subject) {
  return __awaiter(this, void 0, void 0, function () {
    var token, base, candidates, documents, _i, candidates_4, url, res, json, _a, normalizedSearchSubject, _b, documents_1, doc, fileName, docSubject, normalizedDocSubject;
    var _c, _d;
    return __generator(this, function (_e) {
      switch (_e.label) {
        case 0:
          console.log("[findDocumentBySubject] Searching for subject in case", {
            caseId: caseId,
            subject: subject
          });
          return [4 /*yield*/, getToken()];
        case 1:
          token = _e.sent();
          return [4 /*yield*/, resolveApiBaseUrl()];
        case 2:
          base = _e.sent();
          candidates = ["".concat(base, "/cases/").concat(encodeURIComponent(caseId), "/documents"), "".concat(base, "/documents?case_id=").concat(encodeURIComponent(caseId)), "".concat(base, "/cases/").concat(encodeURIComponent(caseId), "/files")];
          documents = [];
          _i = 0, candidates_4 = candidates;
          _e.label = 3;
        case 3:
          if (!(_i < candidates_4.length)) return [3 /*break*/, 9];
          url = candidates_4[_i];
          _e.label = 4;
        case 4:
          _e.trys.push([4, 7,, 8]);
          return [4 /*yield*/, fetch(url, {
            method: "GET",
            headers: {
              Authentication: token,
              "Content-Type": "application/json",
              "Accept-Encoding": "identity"
            }
          })];
        case 5:
          res = _e.sent();
          if (res.status === 404 || res.status === 405) {
            return [3 /*break*/, 8]; // Try next endpoint
          }
          if (!res.ok) {
            return [3 /*break*/, 8];
          }
          return [4 /*yield*/, res.json()];
        case 6:
          json = _e.sent();
          // Handle different response structures
          documents = Array.isArray(json) ? json : Array.isArray(json.documents) ? json.documents : Array.isArray(json.files) ? json.files : Array.isArray(json.items) ? json.items : [];
          if (documents.length >= 0) {
            console.log("[findDocumentBySubject] Found", documents.length, "documents in case");
            return [3 /*break*/, 9]; // Success
          }
          return [3 /*break*/, 8];
        case 7:
          _a = _e.sent();
          return [3 /*break*/, 8];
        case 8:
          _i++;
          return [3 /*break*/, 3];
        case 9:
          if (documents.length === 0) {
            console.log("[findDocumentBySubject] No documents found in case");
            return [2 /*return*/, null];
          }
          normalizedSearchSubject = normalizeSubject(subject);
          console.log("[findDocumentBySubject] Normalized search subject:", normalizedSearchSubject);
          if (!normalizedSearchSubject) {
            console.warn("[findDocumentBySubject] Empty normalized subject, skipping");
            return [2 /*return*/, null];
          }
          // Search for .eml files with matching subject
          for (_b = 0, documents_1 = documents; _b < documents_1.length; _b++) {
            doc = documents_1[_b];
            fileName = String(doc.name || doc.filename || "");
            if (!fileName.toLowerCase().endsWith(".eml")) {
              continue;
            }
            docSubject = ((_c = doc.metadata) === null || _c === void 0 ? void 0 : _c.subject) ||
            // Metadata field (if we stored it)
            doc.subject || (
            // Direct field
            (_d = doc.properties) === null || _d === void 0 ? void 0 : _d.subject) ||
            // Properties object
            "";
            // Fallback: extract from filename (remove .eml extension)
            if (!docSubject) {
              docSubject = fileName.replace(/\.eml$/i, "");
            }
            normalizedDocSubject = normalizeSubject(docSubject);
            console.log("[findDocumentBySubject] Comparing", {
              fileName: fileName,
              docSubject: docSubject,
              normalizedDocSubject: normalizedDocSubject,
              matches: normalizedDocSubject === normalizedSearchSubject
            });
            if (normalizedDocSubject === normalizedSearchSubject) {
              console.log("[findDocumentBySubject] Match found!", {
                id: doc.id,
                name: fileName
              });
              return [2 /*return*/, {
                id: String(doc.id || doc._id),
                name: fileName,
                subject: docSubject
              }];
            }
          }
          console.log("[findDocumentBySubject] No matching document found");
          return [2 /*return*/, null];
      }
    });
  });
}
/**
 * Check if an email with this conversationId and subject is already filed
 * Searches across all cases in the workspace
 *
 * This is the definitive server-side check for "already filed" status
 * Uses conversationId + normalized subject for reliable cross-mailbox matching
 *
 * @param conversationId - Office.js conversationId (available at send time)
 * @param subject - Email subject for additional matching
 * @returns Document info if found, null otherwise
 */
function checkFiledStatusByConversationAndSubject(conversationId, subject) {
  return __awaiter(this, void 0, void 0, function () {
    var normalizedSearchSubject, token, base, listUrl, documents, res, json, e_7, _i, documents_2, doc, fileName, docConversationId, docSubject, normalizedDocSubject, e_8;
    var _a, _b;
    return __generator(this, function (_c) {
      switch (_c.label) {
        case 0:
          if (!conversationId) {
            console.log("[checkFiledStatusByConversationAndSubject] No conversationId provided");
            return [2 /*return*/, null];
          }
          if (!subject) {
            console.log("[checkFiledStatusByConversationAndSubject] No subject provided");
            return [2 /*return*/, null];
          }
          normalizedSearchSubject = normalizeSubject(subject);
          console.log("[checkFiledStatusByConversationAndSubject] Checking:", {
            conversationId: conversationId.substring(0, 30) + "...",
            subject: subject,
            normalizedSubject: normalizedSearchSubject
          });
          _c.label = 1;
        case 1:
          _c.trys.push([1, 11,, 12]);
          return [4 /*yield*/, getToken()];
        case 2:
          token = _c.sent();
          return [4 /*yield*/, resolveApiBaseUrl()];
        case 3:
          base = _c.sent();
          // Search for documents by conversationId in metadata
          // Since backend may not support metadata search, we'll list recent documents and filter manually
          console.log("[checkFiledStatusByConversationAndSubject] Fetching recent documents for matching");
          listUrl = "".concat(base, "/documents?limit=200&sort=-modified_at");
          documents = [];
          _c.label = 4;
        case 4:
          _c.trys.push([4, 9,, 10]);
          return [4 /*yield*/, fetch(listUrl, {
            method: "GET",
            headers: {
              Authentication: token,
              "Content-Type": "application/json",
              "Accept-Encoding": "identity"
            }
          })];
        case 5:
          res = _c.sent();
          if (!res.ok) return [3 /*break*/, 7];
          return [4 /*yield*/, res.json()];
        case 6:
          json = _c.sent();
          documents = Array.isArray(json) ? json : Array.isArray(json.documents) ? json.documents : [];
          console.log("[checkFiledStatusByConversationAndSubject] Fetched", documents.length, "documents for manual search");
          return [3 /*break*/, 8];
        case 7:
          console.warn("[checkFiledStatusByConversationAndSubject] Failed to fetch documents:", res.status);
          return [2 /*return*/, null];
        case 8:
          return [3 /*break*/, 10];
        case 9:
          e_7 = _c.sent();
          console.error("[checkFiledStatusByConversationAndSubject] Fetch failed:", e_7);
          return [2 /*return*/, null];
        case 10:
          // Search through documents for matching conversationId AND normalized subject
          for (_i = 0, documents_2 = documents; _i < documents_2.length; _i++) {
            doc = documents_2[_i];
            fileName = String(doc.name || doc.filename || "");
            if (!fileName.toLowerCase().endsWith(".eml")) {
              continue;
            }
            docConversationId = (_a = doc.metadata) === null || _a === void 0 ? void 0 : _a.conversationId;
            if (!docConversationId || String(docConversationId).trim() !== conversationId.trim()) {
              continue;
            }
            docSubject = ((_b = doc.metadata) === null || _b === void 0 ? void 0 : _b.subject) || doc.subject || "";
            if (!docSubject) {
              // Fallback: extract from filename
              docSubject = fileName.replace(/\.eml$/i, "");
            }
            normalizedDocSubject = normalizeSubject(docSubject);
            if (normalizedDocSubject === normalizedSearchSubject) {
              console.log("[checkFiledStatusByConversationAndSubject] Match found!", {
                documentId: doc.id,
                caseId: doc.case_id,
                subject: docSubject,
                conversationIdMatch: true,
                subjectMatch: true
              });
              return [2 /*return*/, {
                documentId: String(doc.id || doc._id),
                caseId: String(doc.case_id || doc.caseId),
                caseName: doc.case_name || doc.caseName,
                caseKey: doc.case_key || doc.caseKey,
                subject: docSubject
              }];
            }
          }
          console.log("[checkFiledStatusByConversationAndSubject] No match found (checked", documents.length, "documents)");
          return [2 /*return*/, null];
        case 11:
          e_8 = _c.sent();
          console.error("[checkFiledStatusByConversationAndSubject] Error:", e_8);
          return [2 /*return*/, null];
        case 12:
          return [2 /*return*/];
      }
    });
  });
}

/***/ }),

/***/ "./src/utils/constants.ts":
/*!********************************!*\
  !*** ./src/utils/constants.ts ***!
  \********************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   STORAGE_KEYS: function() { return /* binding */ STORAGE_KEYS; }
/* harmony export */ });
var STORAGE_KEYS = {
  onboardingDone: "sc:onboardingDone",
  workspaceId: "sc:workspaceId",
  workspaceName: "sc:workspaceName",
  workspaceHost: "sc:workspaceHost",
  agreementAccepted: "sc:agreementAccepted",
  publicToken: "sc:publicToken",
  recipientHistory: "recipientHistory",
  outlookFolderCache: "sc:outlookFolderCache"
};

/***/ }),

/***/ "./src/utils/filedCache.ts":
/*!*********************************!*\
  !*** ./src/utils/filedCache.ts ***!
  \*********************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   cacheFiledEmail: function() { return /* binding */ cacheFiledEmail; },
/* harmony export */   cacheFiledEmailBySubject: function() { return /* binding */ cacheFiledEmailBySubject; },
/* harmony export */   findFiledEmailBySubject: function() { return /* binding */ findFiledEmailBySubject; },
/* harmony export */   getFiledEmailFromCache: function() { return /* binding */ getFiledEmailFromCache; },
/* harmony export */   removeFiledEmailFromCache: function() { return /* binding */ removeFiledEmailFromCache; }
/* harmony export */ });
/* harmony import */ var _storage__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./storage */ "./src/utils/storage.ts");
/* provided dependency */ var Promise = __webpack_require__(/*! es6-promise */ "./node_modules/es6-promise/dist/es6-promise.js")["Promise"];
var __awaiter = undefined && undefined.__awaiter || function (thisArg, _arguments, P, generator) {
  function adopt(value) {
    return value instanceof P ? value : new P(function (resolve) {
      resolve(value);
    });
  }
  return new (P || (P = Promise))(function (resolve, reject) {
    function fulfilled(value) {
      try {
        step(generator.next(value));
      } catch (e) {
        reject(e);
      }
    }
    function rejected(value) {
      try {
        step(generator["throw"](value));
      } catch (e) {
        reject(e);
      }
    }
    function step(result) {
      result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected);
    }
    step((generator = generator.apply(thisArg, _arguments || [])).next());
  });
};
var __generator = undefined && undefined.__generator || function (thisArg, body) {
  var _ = {
      label: 0,
      sent: function sent() {
        if (t[0] & 1) throw t[1];
        return t[1];
      },
      trys: [],
      ops: []
    },
    f,
    y,
    t,
    g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
  return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function () {
    return this;
  }), g;
  function verb(n) {
    return function (v) {
      return step([n, v]);
    };
  }
  function step(op) {
    if (f) throw new TypeError("Generator is already executing.");
    while (g && (g = 0, op[0] && (_ = 0)), _) try {
      if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
      if (y = 0, t) op = [op[0] & 2, t.value];
      switch (op[0]) {
        case 0:
        case 1:
          t = op;
          break;
        case 4:
          _.label++;
          return {
            value: op[1],
            done: false
          };
        case 5:
          _.label++;
          y = op[1];
          op = [0];
          continue;
        case 7:
          op = _.ops.pop();
          _.trys.pop();
          continue;
        default:
          if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) {
            _ = 0;
            continue;
          }
          if (op[0] === 3 && (!t || op[1] > t[0] && op[1] < t[3])) {
            _.label = op[1];
            break;
          }
          if (op[0] === 6 && _.label < t[1]) {
            _.label = t[1];
            t = op;
            break;
          }
          if (t && _.label < t[2]) {
            _.label = t[2];
            _.ops.push(op);
            break;
          }
          if (t[2]) _.ops.pop();
          _.trys.pop();
          continue;
      }
      op = body.call(thisArg, _);
    } catch (e) {
      op = [6, e];
      y = 0;
    } finally {
      f = t = 0;
    }
    if (op[0] & 5) throw op[1];
    return {
      value: op[0] ? op[1] : void 0,
      done: true
    };
  }
};

var FILED_CACHE_KEY = "sc:filedEmailsCache";
/**
 * Store filed email info by conversationId
 * This enables "already filed" detection for self-sent emails and replies
 *
 * Works for:
 * - Self-sent emails (sender opens received copy)
 * - Sent items (user reopens their own sent email)
 * - Replies in same thread (same conversationId)
 */
function cacheFiledEmail(conversationId, caseId, documentId, subject, caseName, caseKey) {
  return __awaiter(this, void 0, void 0, function () {
    var platform, raw, cache, entries, keep, newCache_1, verification, verifiedCache, writeSuccess, e_1;
    var _a, _b, _c, _d, _e, _f, _g;
    return __generator(this, function (_h) {
      switch (_h.label) {
        case 0:
          if (!conversationId) {
            console.warn("[cacheFiledEmail] No conversationId provided, skipping cache");
            return [2 /*return*/];
          }
          _h.label = 1;
        case 1:
          _h.trys.push([1, 8,, 9]);
          platform = {
            host: (_c = (_b = (_a = Office === null || Office === void 0 ? void 0 : Office.context) === null || _a === void 0 ? void 0 : _a.mailbox) === null || _b === void 0 ? void 0 : _b.diagnostics) === null || _c === void 0 ? void 0 : _c.hostName,
            hostVersion: (_f = (_e = (_d = Office === null || Office === void 0 ? void 0 : Office.context) === null || _d === void 0 ? void 0 : _d.mailbox) === null || _e === void 0 ? void 0 : _e.diagnostics) === null || _f === void 0 ? void 0 : _f.hostVersion,
            platform: (_g = Office === null || Office === void 0 ? void 0 : Office.context) === null || _g === void 0 ? void 0 : _g.platform
          };
          console.log("[cacheFiledEmail] Platform info:", platform);
          return [4 /*yield*/, (0,_storage__WEBPACK_IMPORTED_MODULE_0__.getStored)(FILED_CACHE_KEY)];
        case 2:
          raw = _h.sent();
          cache = raw ? JSON.parse(String(raw)) : {};
          console.log("[cacheFiledEmail] Current cache size:", Object.keys(cache).length);
          cache[conversationId] = {
            caseId: caseId,
            documentId: documentId,
            subject: subject,
            caseName: caseName,
            caseKey: caseKey,
            filedAt: Date.now()
          };
          entries = Object.entries(cache);
          if (!(entries.length > 100)) return [3 /*break*/, 4];
          entries.sort(function (a, b) {
            return b[1].filedAt - a[1].filedAt;
          });
          keep = entries.slice(0, 100);
          newCache_1 = {};
          keep.forEach(function (_a) {
            var key = _a[0],
              val = _a[1];
            newCache_1[key] = val;
          });
          return [4 /*yield*/, (0,_storage__WEBPACK_IMPORTED_MODULE_0__.setStored)(FILED_CACHE_KEY, JSON.stringify(newCache_1))];
        case 3:
          _h.sent();
          console.log("[cacheFiledEmail] Cleaned cache, kept 100 most recent entries");
          return [3 /*break*/, 6];
        case 4:
          return [4 /*yield*/, (0,_storage__WEBPACK_IMPORTED_MODULE_0__.setStored)(FILED_CACHE_KEY, JSON.stringify(cache))];
        case 5:
          _h.sent();
          _h.label = 6;
        case 6:
          return [4 /*yield*/, (0,_storage__WEBPACK_IMPORTED_MODULE_0__.getStored)(FILED_CACHE_KEY)];
        case 7:
          verification = _h.sent();
          verifiedCache = verification ? JSON.parse(String(verification)) : {};
          writeSuccess = !!verifiedCache[conversationId];
          console.log("[cacheFiledEmail] Write verification:", {
            success: writeSuccess,
            cacheSize: Object.keys(verifiedCache).length
          });
          console.log("[cacheFiledEmail] Cached filed email", {
            conversationId: conversationId.substring(0, 20) + "...",
            caseId: caseId,
            documentId: documentId,
            subject: subject,
            writeVerified: writeSuccess
          });
          return [3 /*break*/, 9];
        case 8:
          e_1 = _h.sent();
          console.warn("[cacheFiledEmail] Failed to cache:", e_1);
          return [3 /*break*/, 9];
        case 9:
          return [2 /*return*/];
      }
    });
  });
}
/**
 * Check if email with this conversationId was filed
 * Returns cached info if found, null otherwise
 */
function getFiledEmailFromCache(conversationId) {
  return __awaiter(this, void 0, void 0, function () {
    var platform, raw, cache, cacheKeys, entry, e_2;
    var _a, _b, _c, _d, _e, _f, _g;
    return __generator(this, function (_h) {
      switch (_h.label) {
        case 0:
          if (!conversationId) {
            return [2 /*return*/, null];
          }
          _h.label = 1;
        case 1:
          _h.trys.push([1, 3,, 4]);
          platform = {
            host: (_c = (_b = (_a = Office === null || Office === void 0 ? void 0 : Office.context) === null || _a === void 0 ? void 0 : _a.mailbox) === null || _b === void 0 ? void 0 : _b.diagnostics) === null || _c === void 0 ? void 0 : _c.hostName,
            hostVersion: (_f = (_e = (_d = Office === null || Office === void 0 ? void 0 : Office.context) === null || _d === void 0 ? void 0 : _d.mailbox) === null || _e === void 0 ? void 0 : _e.diagnostics) === null || _f === void 0 ? void 0 : _f.hostVersion,
            platform: (_g = Office === null || Office === void 0 ? void 0 : Office.context) === null || _g === void 0 ? void 0 : _g.platform
          };
          console.log("[getFiledEmailFromCache] Platform info:", platform);
          return [4 /*yield*/, (0,_storage__WEBPACK_IMPORTED_MODULE_0__.getStored)(FILED_CACHE_KEY, true)];
        case 2:
          raw = _h.sent();
          if (!raw) {
            console.log("[getFiledEmailFromCache] No cache found in storage");
            return [2 /*return*/, null];
          }
          cache = JSON.parse(String(raw));
          cacheKeys = Object.keys(cache);
          console.log("[getFiledEmailFromCache] Cache size:", cacheKeys.length, "keys");
          console.log("[getFiledEmailFromCache] Looking for conversationId:", conversationId.substring(0, 30) + "...");
          console.log("[getFiledEmailFromCache] Sample cache keys:", cacheKeys.slice(0, 3).map(function (k) {
            return k.substring(0, 30) + "...";
          }));
          entry = cache[conversationId];
          if (entry) {
            console.log("[getFiledEmailFromCache] ✅ Found cache entry", {
              conversationId: conversationId.substring(0, 20) + "...",
              caseId: entry.caseId,
              documentId: entry.documentId,
              filedAt: new Date(entry.filedAt).toISOString(),
              subject: entry.subject
            });
            return [2 /*return*/, entry];
          }
          console.log("[getFiledEmailFromCache] ❌ No entry for this conversationId");
          return [2 /*return*/, null];
        case 3:
          e_2 = _h.sent();
          console.warn("[getFiledEmailFromCache] Failed to read cache:", e_2);
          return [2 /*return*/, null];
        case 4:
          return [2 /*return*/];
      }
    });
  });
}
/**
 * Cache filed email by subject (fallback when conversationId not available at send time)
 * Used for NEW compose emails where conversationId isn't assigned until after send
 */
function cacheFiledEmailBySubject(subject, caseId, documentId, caseName, caseKey) {
  return __awaiter(this, void 0, void 0, function () {
    var platform, raw, cache, tempKey, entries, keep, newCache_2, verification, verifiedCache, writeSuccess, e_3;
    var _a, _b, _c, _d, _e, _f, _g;
    return __generator(this, function (_h) {
      switch (_h.label) {
        case 0:
          if (!subject) {
            console.warn("[cacheFiledEmailBySubject] No subject provided, skipping cache");
            return [2 /*return*/];
          }
          _h.label = 1;
        case 1:
          _h.trys.push([1, 8,, 9]);
          platform = {
            host: (_c = (_b = (_a = Office === null || Office === void 0 ? void 0 : Office.context) === null || _a === void 0 ? void 0 : _a.mailbox) === null || _b === void 0 ? void 0 : _b.diagnostics) === null || _c === void 0 ? void 0 : _c.hostName,
            hostVersion: (_f = (_e = (_d = Office === null || Office === void 0 ? void 0 : Office.context) === null || _d === void 0 ? void 0 : _d.mailbox) === null || _e === void 0 ? void 0 : _e.diagnostics) === null || _f === void 0 ? void 0 : _f.hostVersion,
            platform: (_g = Office === null || Office === void 0 ? void 0 : Office.context) === null || _g === void 0 ? void 0 : _g.platform
          };
          console.log("[cacheFiledEmailBySubject] Platform info:", platform);
          return [4 /*yield*/, (0,_storage__WEBPACK_IMPORTED_MODULE_0__.getStored)(FILED_CACHE_KEY)];
        case 2:
          raw = _h.sent();
          cache = raw ? JSON.parse(String(raw)) : {};
          console.log("[cacheFiledEmailBySubject] Current cache size:", Object.keys(cache).length);
          tempKey = "subj:".concat(subject.trim().toLowerCase());
          console.log("[cacheFiledEmailBySubject] Using temp key:", tempKey);
          cache[tempKey] = {
            caseId: caseId,
            documentId: documentId,
            subject: subject,
            caseName: caseName,
            caseKey: caseKey,
            filedAt: Date.now()
          };
          entries = Object.entries(cache);
          if (!(entries.length > 100)) return [3 /*break*/, 4];
          entries.sort(function (a, b) {
            return b[1].filedAt - a[1].filedAt;
          });
          keep = entries.slice(0, 100);
          newCache_2 = {};
          keep.forEach(function (_a) {
            var key = _a[0],
              val = _a[1];
            newCache_2[key] = val;
          });
          return [4 /*yield*/, (0,_storage__WEBPACK_IMPORTED_MODULE_0__.setStored)(FILED_CACHE_KEY, JSON.stringify(newCache_2))];
        case 3:
          _h.sent();
          console.log("[cacheFiledEmailBySubject] Cleaned cache, kept 100 most recent entries");
          return [3 /*break*/, 6];
        case 4:
          return [4 /*yield*/, (0,_storage__WEBPACK_IMPORTED_MODULE_0__.setStored)(FILED_CACHE_KEY, JSON.stringify(cache))];
        case 5:
          _h.sent();
          _h.label = 6;
        case 6:
          return [4 /*yield*/, (0,_storage__WEBPACK_IMPORTED_MODULE_0__.getStored)(FILED_CACHE_KEY)];
        case 7:
          verification = _h.sent();
          verifiedCache = verification ? JSON.parse(String(verification)) : {};
          writeSuccess = !!verifiedCache[tempKey];
          console.log("[cacheFiledEmailBySubject] Write verification:", {
            success: writeSuccess,
            cacheSize: Object.keys(verifiedCache).length,
            tempKey: tempKey
          });
          console.log("[cacheFiledEmailBySubject] Cached filed email by subject", {
            subject: subject,
            caseId: caseId,
            documentId: documentId,
            writeVerified: writeSuccess
          });
          return [3 /*break*/, 9];
        case 8:
          e_3 = _h.sent();
          console.warn("[cacheFiledEmailBySubject] Failed to cache:", e_3);
          return [3 /*break*/, 9];
        case 9:
          return [2 /*return*/];
      }
    });
  });
}
/**
 * Search cache by subject (fallback when conversationId lookup fails)
 * Also upgrades the cache entry to use conversationId for future lookups
 */
function findFiledEmailBySubject(subject, conversationId) {
  return __awaiter(this, void 0, void 0, function () {
    var platform, raw, cache, cacheKeys, tempKey, entry, verification, verifiedCache, upgradeSuccess, e_4;
    var _a, _b, _c, _d, _e, _f, _g;
    return __generator(this, function (_h) {
      switch (_h.label) {
        case 0:
          if (!subject) {
            return [2 /*return*/, null];
          }
          _h.label = 1;
        case 1:
          _h.trys.push([1, 7,, 8]);
          platform = {
            host: (_c = (_b = (_a = Office === null || Office === void 0 ? void 0 : Office.context) === null || _a === void 0 ? void 0 : _a.mailbox) === null || _b === void 0 ? void 0 : _b.diagnostics) === null || _c === void 0 ? void 0 : _c.hostName,
            hostVersion: (_f = (_e = (_d = Office === null || Office === void 0 ? void 0 : Office.context) === null || _d === void 0 ? void 0 : _d.mailbox) === null || _e === void 0 ? void 0 : _e.diagnostics) === null || _f === void 0 ? void 0 : _f.hostVersion,
            platform: (_g = Office === null || Office === void 0 ? void 0 : Office.context) === null || _g === void 0 ? void 0 : _g.platform
          };
          console.log("[findFiledEmailBySubject] Platform info:", platform);
          return [4 /*yield*/, (0,_storage__WEBPACK_IMPORTED_MODULE_0__.getStored)(FILED_CACHE_KEY, true)];
        case 2:
          raw = _h.sent();
          if (!raw) {
            console.log("[findFiledEmailBySubject] No cache found in storage");
            return [2 /*return*/, null];
          }
          cache = JSON.parse(String(raw));
          cacheKeys = Object.keys(cache);
          console.log("[findFiledEmailBySubject] Cache size:", cacheKeys.length, "keys");
          tempKey = "subj:".concat(subject.trim().toLowerCase());
          console.log("[findFiledEmailBySubject] Looking for temp key:", tempKey);
          console.log("[findFiledEmailBySubject] Subject-based keys in cache:", cacheKeys.filter(function (k) {
            return k.startsWith("subj:");
          }).length);
          entry = cache[tempKey];
          if (!entry) return [3 /*break*/, 6];
          console.log("[findFiledEmailBySubject] ✅ Found cache entry by subject", {
            subject: subject,
            caseId: entry.caseId,
            documentId: entry.documentId,
            filedAt: new Date(entry.filedAt).toISOString()
          });
          if (!conversationId) return [3 /*break*/, 5];
          console.log("[findFiledEmailBySubject] Upgrading cache with conversationId:", conversationId.substring(0, 30) + "...");
          cache[conversationId] = entry;
          // Keep the subject-based entry for a while (don't delete)
          return [4 /*yield*/, (0,_storage__WEBPACK_IMPORTED_MODULE_0__.setStored)(FILED_CACHE_KEY, JSON.stringify(cache))];
        case 3:
          // Keep the subject-based entry for a while (don't delete)
          _h.sent();
          return [4 /*yield*/, (0,_storage__WEBPACK_IMPORTED_MODULE_0__.getStored)(FILED_CACHE_KEY)];
        case 4:
          verification = _h.sent();
          verifiedCache = verification ? JSON.parse(String(verification)) : {};
          upgradeSuccess = !!verifiedCache[conversationId];
          console.log("[findFiledEmailBySubject] Cache upgrade verification:", {
            success: upgradeSuccess,
            cacheSize: Object.keys(verifiedCache).length
          });
          _h.label = 5;
        case 5:
          return [2 /*return*/, entry];
        case 6:
          console.log("[findFiledEmailBySubject] ❌ No entry for this subject");
          return [2 /*return*/, null];
        case 7:
          e_4 = _h.sent();
          console.warn("[findFiledEmailBySubject] Failed:", e_4);
          return [2 /*return*/, null];
        case 8:
          return [2 /*return*/];
      }
    });
  });
}
/**
 * Remove filed email from cache (e.g., if document was deleted)
 */
function removeFiledEmailFromCache(conversationId) {
  return __awaiter(this, void 0, void 0, function () {
    var raw, cache, e_5;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          if (!conversationId) {
            return [2 /*return*/];
          }
          _a.label = 1;
        case 1:
          _a.trys.push([1, 4,, 5]);
          return [4 /*yield*/, (0,_storage__WEBPACK_IMPORTED_MODULE_0__.getStored)(FILED_CACHE_KEY)];
        case 2:
          raw = _a.sent();
          if (!raw) return [2 /*return*/];
          cache = JSON.parse(String(raw));
          delete cache[conversationId];
          return [4 /*yield*/, (0,_storage__WEBPACK_IMPORTED_MODULE_0__.setStored)(FILED_CACHE_KEY, JSON.stringify(cache))];
        case 3:
          _a.sent();
          console.log("[removeFiledEmailFromCache] Removed entry", {
            conversationId: conversationId.substring(0, 20) + "..."
          });
          return [3 /*break*/, 5];
        case 4:
          e_5 = _a.sent();
          console.warn("[removeFiledEmailFromCache] Failed:", e_5);
          return [3 /*break*/, 5];
        case 5:
          return [2 /*return*/];
      }
    });
  });
}

/***/ }),

/***/ "./src/utils/storage.ts":
/*!******************************!*\
  !*** ./src/utils/storage.ts ***!
  \******************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   clearDebugLog: function() { return /* binding */ clearDebugLog; },
/* harmony export */   getDebugLog: function() { return /* binding */ getDebugLog; },
/* harmony export */   getStored: function() { return /* binding */ getStored; },
/* harmony export */   removeStored: function() { return /* binding */ removeStored; },
/* harmony export */   setStored: function() { return /* binding */ setStored; }
/* harmony export */ });
/* provided dependency */ var Promise = __webpack_require__(/*! es6-promise */ "./node_modules/es6-promise/dist/es6-promise.js")["Promise"];
// src/utils/storage.ts
/* global Office, OfficeRuntime */
var __awaiter = undefined && undefined.__awaiter || function (thisArg, _arguments, P, generator) {
  function adopt(value) {
    return value instanceof P ? value : new P(function (resolve) {
      resolve(value);
    });
  }
  return new (P || (P = Promise))(function (resolve, reject) {
    function fulfilled(value) {
      try {
        step(generator.next(value));
      } catch (e) {
        reject(e);
      }
    }
    function rejected(value) {
      try {
        step(generator["throw"](value));
      } catch (e) {
        reject(e);
      }
    }
    function step(result) {
      result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected);
    }
    step((generator = generator.apply(thisArg, _arguments || [])).next());
  });
};
var __generator = undefined && undefined.__generator || function (thisArg, body) {
  var _ = {
      label: 0,
      sent: function sent() {
        if (t[0] & 1) throw t[1];
        return t[1];
      },
      trys: [],
      ops: []
    },
    f,
    y,
    t,
    g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
  return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function () {
    return this;
  }), g;
  function verb(n) {
    return function (v) {
      return step([n, v]);
    };
  }
  function step(op) {
    if (f) throw new TypeError("Generator is already executing.");
    while (g && (g = 0, op[0] && (_ = 0)), _) try {
      if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
      if (y = 0, t) op = [op[0] & 2, t.value];
      switch (op[0]) {
        case 0:
        case 1:
          t = op;
          break;
        case 4:
          _.label++;
          return {
            value: op[1],
            done: false
          };
        case 5:
          _.label++;
          y = op[1];
          op = [0];
          continue;
        case 7:
          op = _.ops.pop();
          _.trys.pop();
          continue;
        default:
          if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) {
            _ = 0;
            continue;
          }
          if (op[0] === 3 && (!t || op[1] > t[0] && op[1] < t[3])) {
            _.label = op[1];
            break;
          }
          if (op[0] === 6 && _.label < t[1]) {
            _.label = t[1];
            t = op;
            break;
          }
          if (t && _.label < t[2]) {
            _.label = t[2];
            _.ops.push(op);
            break;
          }
          if (t[2]) _.ops.pop();
          _.trys.pop();
          continue;
      }
      op = body.call(thisArg, _);
    } catch (e) {
      op = [6, e];
      y = 0;
    } finally {
      f = t = 0;
    }
    if (op[0] & 5) throw op[1];
    return {
      value: op[0] ? op[1] : void 0,
      done: true
    };
  }
};
// Feature flag: Set to false to silence verbose logging (helps with render loops)
var VERBOSE_LOGGING = false;
// Debug log that persists across sessions
var DEBUG_LOG_KEY = "sc:debugLog";
function getDebugLog() {
  return __awaiter(this, void 0, void 0, function () {
    return __generator(this, function (_a) {
      try {
        if (hasRoamingSettings()) {
          return [2 /*return*/, String(Office.context.roamingSettings.get(DEBUG_LOG_KEY) || "")];
        } else if (typeof localStorage !== "undefined") {
          return [2 /*return*/, localStorage.getItem(DEBUG_LOG_KEY) || ""];
        }
        return [2 /*return*/, ""];
      } catch (_b) {
        return [2 /*return*/, ""];
      }
      // removed by dead control flow

    });
  });
}
function clearDebugLog() {
  return __awaiter(this, void 0, void 0, function () {
    var _a;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          _b.trys.push([0, 4,, 5]);
          if (!hasRoamingSettings()) return [3 /*break*/, 2];
          Office.context.roamingSettings.remove(DEBUG_LOG_KEY);
          return [4 /*yield*/, saveRoamingSettings()];
        case 1:
          _b.sent();
          return [3 /*break*/, 3];
        case 2:
          if (typeof localStorage !== "undefined") {
            localStorage.removeItem(DEBUG_LOG_KEY);
          }
          _b.label = 3;
        case 3:
          return [3 /*break*/, 5];
        case 4:
          _a = _b.sent();
          return [3 /*break*/, 5];
        case 5:
          return [2 /*return*/];
      }
    });
  });
}
function hasOfficeRuntimeStorage() {
  try {
    return typeof OfficeRuntime !== "undefined" && !!(OfficeRuntime === null || OfficeRuntime === void 0 ? void 0 : OfficeRuntime.storage);
  } catch (_a) {
    return false;
  }
}
function hasRoamingSettings() {
  var _a;
  try {
    return !!((_a = Office === null || Office === void 0 ? void 0 : Office.context) === null || _a === void 0 ? void 0 : _a.roamingSettings);
  } catch (_b) {
    return false;
  }
}
function saveRoamingSettings() {
  return __awaiter(this, void 0, void 0, function () {
    var startTime;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          startTime = Date.now();
          if (VERBOSE_LOGGING) console.log("[saveRoamingSettings] Starting saveAsync...");
          return [4 /*yield*/, new Promise(function (resolve, reject) {
            try {
              Office.context.roamingSettings.saveAsync(function (res) {
                var _a, _b, _c;
                var duration = Date.now() - startTime;
                if ((res === null || res === void 0 ? void 0 : res.status) === Office.AsyncResultStatus.Succeeded) {
                  if (VERBOSE_LOGGING) console.log("[saveRoamingSettings] \u2705 Succeeded in ".concat(duration, "ms"));
                  resolve();
                } else {
                  var errorMsg = ((_a = res === null || res === void 0 ? void 0 : res.error) === null || _a === void 0 ? void 0 : _a.message) || "roamingSettings.saveAsync failed";
                  console.error("[saveRoamingSettings] \u274C Failed in ".concat(duration, "ms:"), errorMsg, {
                    status: res === null || res === void 0 ? void 0 : res.status,
                    errorCode: (_b = res === null || res === void 0 ? void 0 : res.error) === null || _b === void 0 ? void 0 : _b.code,
                    errorName: (_c = res === null || res === void 0 ? void 0 : res.error) === null || _c === void 0 ? void 0 : _c.name
                  });
                  reject(new Error(errorMsg));
                }
              });
            } catch (e) {
              var duration = Date.now() - startTime;
              console.error("[saveRoamingSettings] \u274C Exception in ".concat(duration, "ms:"), e);
              reject(e);
            }
          })];
        case 1:
          _a.sent();
          // CRITICAL FOR DESKTOP OUTLOOK: Add small delay to ensure operation completes
          // Desktop Outlook may close compose window immediately after send,
          // interrupting async operations. This delay ensures saveAsync completes.
          return [4 /*yield*/, new Promise(function (resolve) {
            return setTimeout(resolve, 100);
          })];
        case 2:
          // CRITICAL FOR DESKTOP OUTLOOK: Add small delay to ensure operation completes
          // Desktop Outlook may close compose window immediately after send,
          // interrupting async operations. This delay ensures saveAsync completes.
          _a.sent();
          return [2 /*return*/];
      }
    });
  });
}
function getStored(key_1) {
  return __awaiter(this, arguments, void 0, function (key, forceFresh) {
    var k, storageBackend, shouldLog, v_1, v_2, v, e_1;
    if (forceFresh === void 0) {
      forceFresh = false;
    }
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          k = String(key || "").trim();
          if (!k) return [2 /*return*/, null];
          storageBackend = hasOfficeRuntimeStorage() ? "OfficeRuntime.storage" : hasRoamingSettings() ? "roamingSettings" : "localStorage";
          shouldLog = VERBOSE_LOGGING && !k.includes("recipientHistory") && !k.includes("recentCases");
          if (shouldLog) {
            console.log("[getStored] Using storage backend:", storageBackend, "for key:", k, forceFresh ? "(force fresh)" : "");
          }
          _a.label = 1;
        case 1:
          _a.trys.push([1, 7,, 8]);
          if (!hasOfficeRuntimeStorage()) return [3 /*break*/, 3];
          return [4 /*yield*/, OfficeRuntime.storage.getItem(k)];
        case 2:
          v_1 = _a.sent();
          if (shouldLog) {
            console.log("[getStored] Got from OfficeRuntime.storage:", k, v_1 ? "found (".concat(v_1.length, " chars)") : "not found");
          }
          return [2 /*return*/, typeof v_1 === "string" ? v_1 : null];
        case 3:
          if (!hasRoamingSettings()) return [3 /*break*/, 6];
          if (!(forceFresh && VERBOSE_LOGGING)) return [3 /*break*/, 5];
          console.log("[getStored] Waiting 500ms for roamingSettings sync...");
          return [4 /*yield*/, new Promise(function (resolve) {
            return setTimeout(resolve, 500);
          })];
        case 4:
          _a.sent();
          _a.label = 5;
        case 5:
          v_2 = Office.context.roamingSettings.get(k);
          if (shouldLog) {
            console.log("[getStored] Got from roamingSettings:", k, v_2 ? "found (".concat(String(v_2).length, " chars)") : "not found");
          }
          return [2 /*return*/, typeof v_2 === "string" ? v_2 : null];
        case 6:
          v = localStorage.getItem(k);
          if (shouldLog) {
            console.warn("[getStored] No Office storage, using localStorage:", k, v ? "found (".concat(v.length, " chars)") : "not found");
          }
          return [2 /*return*/, v];
        case 7:
          e_1 = _a.sent();
          console.warn("[getStored] Failed, falling back to localStorage:", e_1);
          return [2 /*return*/, localStorage.getItem(k)];
        case 8:
          return [2 /*return*/];
      }
    });
  });
}
function setStored(key_1, value_1) {
  return __awaiter(this, arguments, void 0, function (key, value, retryCount) {
    var k, v, MAX_RETRIES, storageBackend, saveError_1, delay_1, e_2;
    if (retryCount === void 0) {
      retryCount = 0;
    }
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          k = String(key || "").trim();
          if (!k) return [2 /*return*/];
          v = String(value !== null && value !== void 0 ? value : "");
          MAX_RETRIES = 2;
          storageBackend = hasOfficeRuntimeStorage() ? "OfficeRuntime.storage" : hasRoamingSettings() ? "roamingSettings" : "localStorage";
          if (VERBOSE_LOGGING) {
            console.log("[setStored] Using storage backend:", storageBackend, "for key:", k, "(".concat(v.length, " chars)"), retryCount > 0 ? "[retry ".concat(retryCount, "]") : "");
          }
          _a.label = 1;
        case 1:
          _a.trys.push([1, 10,, 11]);
          if (!hasOfficeRuntimeStorage()) return [3 /*break*/, 3];
          if (VERBOSE_LOGGING) console.log("[setStored] Writing to OfficeRuntime.storage...");
          return [4 /*yield*/, OfficeRuntime.storage.setItem(k, v)];
        case 2:
          _a.sent();
          if (VERBOSE_LOGGING) console.log("[setStored] ✅ Write to OfficeRuntime.storage completed");
          return [2 /*return*/];
        case 3:
          if (!hasRoamingSettings()) return [3 /*break*/, 9];
          if (VERBOSE_LOGGING) console.log("[setStored] Writing to roamingSettings...");
          Office.context.roamingSettings.set(k, v);
          if (VERBOSE_LOGGING) console.log("[setStored] Calling saveAsync...");
          _a.label = 4;
        case 4:
          _a.trys.push([4, 6,, 9]);
          return [4 /*yield*/, saveRoamingSettings()];
        case 5:
          _a.sent();
          if (VERBOSE_LOGGING) console.log("[setStored] ✅ saveAsync completed");
          return [2 /*return*/];
        case 6:
          saveError_1 = _a.sent();
          console.error("[setStored] saveAsync failed:", saveError_1);
          if (!(retryCount < MAX_RETRIES)) return [3 /*break*/, 8];
          delay_1 = 200 * (retryCount + 1);
          if (VERBOSE_LOGGING) console.log("[setStored] Retrying in ".concat(delay_1, "ms..."));
          return [4 /*yield*/, new Promise(function (resolve) {
            return setTimeout(resolve, delay_1);
          })];
        case 7:
          _a.sent();
          return [2 /*return*/, setStored(key, value, retryCount + 1)];
        case 8:
          throw saveError_1;
        case 9:
          if (VERBOSE_LOGGING) console.warn("[setStored] No Office storage, using localStorage for key:", k);
          localStorage.setItem(k, v);
          return [3 /*break*/, 11];
        case 10:
          e_2 = _a.sent();
          console.warn("[setStored] ❌ Failed after retries, falling back to localStorage:", e_2);
          localStorage.setItem(k, v);
          return [3 /*break*/, 11];
        case 11:
          return [2 /*return*/];
      }
    });
  });
}
function removeStored(key) {
  return __awaiter(this, void 0, void 0, function () {
    var k, _a;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          k = String(key || "").trim();
          if (!k) return [2 /*return*/];
          _b.label = 1;
        case 1:
          _b.trys.push([1, 6,, 7]);
          if (!hasOfficeRuntimeStorage()) return [3 /*break*/, 3];
          return [4 /*yield*/, OfficeRuntime.storage.removeItem(k)];
        case 2:
          _b.sent();
          return [2 /*return*/];
        case 3:
          if (!hasRoamingSettings()) return [3 /*break*/, 5];
          Office.context.roamingSettings.remove(k);
          return [4 /*yield*/, saveRoamingSettings()];
        case 4:
          _b.sent();
          return [2 /*return*/];
        case 5:
          localStorage.removeItem(k);
          return [3 /*break*/, 7];
        case 6:
          _a = _b.sent();
          localStorage.removeItem(k);
          return [3 /*break*/, 7];
        case 7:
          return [2 /*return*/];
      }
    });
  });
}

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Check if module exists (development only)
/******/ 		if (__webpack_modules__[moduleId] === undefined) {
/******/ 			var e = new Error("Cannot find module '" + moduleId + "'");
/******/ 			e.code = 'MODULE_NOT_FOUND';
/******/ 			throw e;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/define property getters */
/******/ 	!function() {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = function(exports, definition) {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be in strict mode.
!function() {
"use strict";
/*!**********************************!*\
  !*** ./src/commands/commands.ts ***!
  \**********************************/
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _onMessageSendHandler__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./onMessageSendHandler */ "./src/commands/onMessageSendHandler.ts");
/* provided dependency */ var Promise = __webpack_require__(/*! es6-promise */ "./node_modules/es6-promise/dist/es6-promise.js")["Promise"];
/* global Office */
var __awaiter = undefined && undefined.__awaiter || function (thisArg, _arguments, P, generator) {
  function adopt(value) {
    return value instanceof P ? value : new P(function (resolve) {
      resolve(value);
    });
  }
  return new (P || (P = Promise))(function (resolve, reject) {
    function fulfilled(value) {
      try {
        step(generator.next(value));
      } catch (e) {
        reject(e);
      }
    }
    function rejected(value) {
      try {
        step(generator["throw"](value));
      } catch (e) {
        reject(e);
      }
    }
    function step(result) {
      result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected);
    }
    step((generator = generator.apply(thisArg, _arguments || [])).next());
  });
};
var __generator = undefined && undefined.__generator || function (thisArg, body) {
  var _ = {
      label: 0,
      sent: function sent() {
        if (t[0] & 1) throw t[1];
        return t[1];
      },
      trys: [],
      ops: []
    },
    f,
    y,
    t,
    g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
  return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function () {
    return this;
  }), g;
  function verb(n) {
    return function (v) {
      return step([n, v]);
    };
  }
  function step(op) {
    if (f) throw new TypeError("Generator is already executing.");
    while (g && (g = 0, op[0] && (_ = 0)), _) try {
      if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
      if (y = 0, t) op = [op[0] & 2, t.value];
      switch (op[0]) {
        case 0:
        case 1:
          t = op;
          break;
        case 4:
          _.label++;
          return {
            value: op[1],
            done: false
          };
        case 5:
          _.label++;
          y = op[1];
          op = [0];
          continue;
        case 7:
          op = _.ops.pop();
          _.trys.pop();
          continue;
        default:
          if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) {
            _ = 0;
            continue;
          }
          if (op[0] === 3 && (!t || op[1] > t[0] && op[1] < t[3])) {
            _.label = op[1];
            break;
          }
          if (op[0] === 6 && _.label < t[1]) {
            _.label = t[1];
            t = op;
            break;
          }
          if (t && _.label < t[2]) {
            _.label = t[2];
            _.ops.push(op);
            break;
          }
          if (t[2]) _.ops.pop();
          _.trys.pop();
          continue;
      }
      op = body.call(thisArg, _);
    } catch (e) {
      op = [6, e];
      y = 0;
    } finally {
      f = t = 0;
    }
    if (op[0] & 5) throw op[1];
    return {
      value: op[0] ? op[1] : void 0,
      done: true
    };
  }
};

console.log("[commands.ts] Script loaded");
var associated = false;
function associateHandlers() {
  var _a;
  if (associated) return;
  associated = true;
  try {
    if (!((_a = Office === null || Office === void 0 ? void 0 : Office.actions) === null || _a === void 0 ? void 0 : _a.associate)) {
      console.warn("[commands.ts] Office.actions.associate not available");
      return;
    }
    console.log("[commands.ts] Associating onMessageSendHandler");
    Office.actions.associate("onMessageSendHandler", _onMessageSendHandler__WEBPACK_IMPORTED_MODULE_0__.onMessageSendHandler);
    console.log("[commands.ts] Handler associated successfully");
  } catch (e) {
    console.error("[commands.ts] Failed to associate handler:", e);
  }
}
function boot() {
  return __awaiter(this, void 0, void 0, function () {
    var e_1;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          _a.trys.push([0, 4, 5, 6]);
          if (!(typeof (Office === null || Office === void 0 ? void 0 : Office.onReady) === "function")) return [3 /*break*/, 2];
          return [4 /*yield*/, Office.onReady()];
        case 1:
          _a.sent();
          console.log("[commands.ts] Office.onReady fired");
          console.log("[commands.ts] Office.context:", Office.context);
          return [3 /*break*/, 3];
        case 2:
          console.warn("[commands.ts] Office.onReady not available");
          _a.label = 3;
        case 3:
          return [3 /*break*/, 6];
        case 4:
          e_1 = _a.sent();
          console.error("[commands.ts] Office.onReady failed:", e_1);
          return [3 /*break*/, 6];
        case 5:
          associateHandlers();
          return [7 /*endfinally*/];
        case 6:
          return [2 /*return*/];
      }
    });
  });
}
// Start immediately, but also try again onReady.
// This avoids cases where the script runs before Office runtime is fully initialised.
boot();
try {
  if (typeof (Office === null || Office === void 0 ? void 0 : Office.onReady) === "function") {
    Office.onReady(function () {
      associateHandlers();
    });
  }
} catch (_a) {
  // ignore
}
}();
/******/ })()
;
//# sourceMappingURL=commands.js.map