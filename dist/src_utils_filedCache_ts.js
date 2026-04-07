"use strict";
(self["webpackChunkoffice_addin_taskpane_react"] = self["webpackChunkoffice_addin_taskpane_react"] || []).push([["src_utils_filedCache_ts"],{

/***/ "./src/utils/filedCache.ts":
/*!*********************************!*\
  !*** ./src/utils/filedCache.ts ***!
  \*********************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   cacheFiledEmail: function() { return /* binding */ cacheFiledEmail; },
/* harmony export */   cacheFiledEmailBySubject: function() { return /* binding */ cacheFiledEmailBySubject; },
/* harmony export */   findFiledEmailBySubject: function() { return /* binding */ findFiledEmailBySubject; },
/* harmony export */   getFiledEmailFromCache: function() { return /* binding */ getFiledEmailFromCache; },
/* harmony export */   removeFiledEmailFromCache: function() { return /* binding */ removeFiledEmailFromCache; }
/* harmony export */ });
/* provided dependency */ var Promise = __webpack_require__(/*! es6-promise */ "./node_modules/es6-promise/dist/es6-promise.js")["Promise"];
// filedCache.ts uses localStorage directly (not setStored/getStored) because:
// - Filed email cache is per-device duplicate detection state — no cross-device sync needed.
// - setStored falls back to roamingSettings in OWA (OfficeRuntime.storage is Desktop-only),
//   and writing sc:filedEmailsCache to roamingSettings contributes to the 32KB overflow.
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
function lsGet() {
  try {
    return typeof localStorage !== "undefined" ? localStorage.getItem(FILED_CACHE_KEY) : null;
  } catch (_a) {
    return null;
  }
}
function lsSet(value) {
  try {
    if (typeof localStorage !== "undefined") {
      localStorage.setItem(FILED_CACHE_KEY, value);
    }
  } catch (_a) {
    // localStorage full or unavailable — silently ignore, cache is non-critical
  }
}
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
    var platform, raw, cache, entries, keep, newCache_1, verification, verifiedCache, writeSuccess;
    var _a, _b, _c, _d, _e, _f, _g;
    return __generator(this, function (_h) {
      if (!conversationId) {
        console.warn("[cacheFiledEmail] No conversationId provided, skipping cache");
        return [2 /*return*/];
      }
      try {
        platform = {
          host: (_c = (_b = (_a = Office === null || Office === void 0 ? void 0 : Office.context) === null || _a === void 0 ? void 0 : _a.mailbox) === null || _b === void 0 ? void 0 : _b.diagnostics) === null || _c === void 0 ? void 0 : _c.hostName,
          hostVersion: (_f = (_e = (_d = Office === null || Office === void 0 ? void 0 : Office.context) === null || _d === void 0 ? void 0 : _d.mailbox) === null || _e === void 0 ? void 0 : _e.diagnostics) === null || _f === void 0 ? void 0 : _f.hostVersion,
          platform: (_g = Office === null || Office === void 0 ? void 0 : Office.context) === null || _g === void 0 ? void 0 : _g.platform
        };
        console.log("[cacheFiledEmail] Platform info:", platform);
        raw = lsGet();
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
        entries.sort(function (a, b) {
          return b[1].filedAt - a[1].filedAt;
        });
        keep = entries.slice(0, 8);
        newCache_1 = {};
        keep.forEach(function (_a) {
          var key = _a[0],
            val = _a[1];
          newCache_1[key] = val;
        });
        lsSet(JSON.stringify(newCache_1));
        if (entries.length > 8) {
          console.log("[cacheFiledEmail] Pruned cache from", entries.length, "to 8 entries");
        }
        verification = lsGet();
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
      } catch (e) {
        console.warn("[cacheFiledEmail] Failed to cache:", e);
        // Non-critical, don't throw
      }
      return [2 /*return*/];
    });
  });
}
/**
 * Check if email with this conversationId was filed
 * Returns cached info if found, null otherwise
 */
function getFiledEmailFromCache(conversationId) {
  return __awaiter(this, void 0, void 0, function () {
    var platform, raw, cache, cacheKeys, entry;
    var _a, _b, _c, _d, _e, _f, _g;
    return __generator(this, function (_h) {
      if (!conversationId) {
        return [2 /*return*/, null];
      }
      try {
        platform = {
          host: (_c = (_b = (_a = Office === null || Office === void 0 ? void 0 : Office.context) === null || _a === void 0 ? void 0 : _a.mailbox) === null || _b === void 0 ? void 0 : _b.diagnostics) === null || _c === void 0 ? void 0 : _c.hostName,
          hostVersion: (_f = (_e = (_d = Office === null || Office === void 0 ? void 0 : Office.context) === null || _d === void 0 ? void 0 : _d.mailbox) === null || _e === void 0 ? void 0 : _e.diagnostics) === null || _f === void 0 ? void 0 : _f.hostVersion,
          platform: (_g = Office === null || Office === void 0 ? void 0 : Office.context) === null || _g === void 0 ? void 0 : _g.platform
        };
        console.log("[getFiledEmailFromCache] Platform info:", platform);
        raw = lsGet();
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
      } catch (e) {
        console.warn("[getFiledEmailFromCache] Failed to read cache:", e);
        return [2 /*return*/, null];
      }
      // removed by dead control flow

    });
  });
}
/**
 * Cache filed email by subject (fallback when conversationId not available at send time)
 * Used for NEW compose emails where conversationId isn't assigned until after send
 */
function cacheFiledEmailBySubject(subject, caseId, documentId, caseName, caseKey) {
  return __awaiter(this, void 0, void 0, function () {
    var platform, raw, cache, tempKey, entries, keep, newCache_2, verification, verifiedCache, writeSuccess;
    var _a, _b, _c, _d, _e, _f, _g;
    return __generator(this, function (_h) {
      if (!subject) {
        console.warn("[cacheFiledEmailBySubject] No subject provided, skipping cache");
        return [2 /*return*/];
      }
      try {
        platform = {
          host: (_c = (_b = (_a = Office === null || Office === void 0 ? void 0 : Office.context) === null || _a === void 0 ? void 0 : _a.mailbox) === null || _b === void 0 ? void 0 : _b.diagnostics) === null || _c === void 0 ? void 0 : _c.hostName,
          hostVersion: (_f = (_e = (_d = Office === null || Office === void 0 ? void 0 : Office.context) === null || _d === void 0 ? void 0 : _d.mailbox) === null || _e === void 0 ? void 0 : _e.diagnostics) === null || _f === void 0 ? void 0 : _f.hostVersion,
          platform: (_g = Office === null || Office === void 0 ? void 0 : Office.context) === null || _g === void 0 ? void 0 : _g.platform
        };
        console.log("[cacheFiledEmailBySubject] Platform info:", platform);
        raw = lsGet();
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
        entries.sort(function (a, b) {
          return b[1].filedAt - a[1].filedAt;
        });
        keep = entries.slice(0, 8);
        newCache_2 = {};
        keep.forEach(function (_a) {
          var key = _a[0],
            val = _a[1];
          newCache_2[key] = val;
        });
        lsSet(JSON.stringify(newCache_2));
        if (entries.length > 8) {
          console.log("[cacheFiledEmailBySubject] Pruned cache from", entries.length, "to 8 entries");
        }
        verification = lsGet();
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
      } catch (e) {
        console.warn("[cacheFiledEmailBySubject] Failed to cache:", e);
      }
      return [2 /*return*/];
    });
  });
}
/**
 * Search cache by subject (fallback when conversationId lookup fails)
 * Also upgrades the cache entry to use conversationId for future lookups
 */
function findFiledEmailBySubject(subject, conversationId) {
  return __awaiter(this, void 0, void 0, function () {
    var platform, raw, cache, cacheKeys, tempKey, entry, verification, verifiedCache, upgradeSuccess;
    var _a, _b, _c, _d, _e, _f, _g;
    return __generator(this, function (_h) {
      if (!subject) {
        return [2 /*return*/, null];
      }
      try {
        platform = {
          host: (_c = (_b = (_a = Office === null || Office === void 0 ? void 0 : Office.context) === null || _a === void 0 ? void 0 : _a.mailbox) === null || _b === void 0 ? void 0 : _b.diagnostics) === null || _c === void 0 ? void 0 : _c.hostName,
          hostVersion: (_f = (_e = (_d = Office === null || Office === void 0 ? void 0 : Office.context) === null || _d === void 0 ? void 0 : _d.mailbox) === null || _e === void 0 ? void 0 : _e.diagnostics) === null || _f === void 0 ? void 0 : _f.hostVersion,
          platform: (_g = Office === null || Office === void 0 ? void 0 : Office.context) === null || _g === void 0 ? void 0 : _g.platform
        };
        console.log("[findFiledEmailBySubject] Platform info:", platform);
        raw = lsGet();
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
        if (entry) {
          console.log("[findFiledEmailBySubject] ✅ Found cache entry by subject", {
            subject: subject,
            caseId: entry.caseId,
            documentId: entry.documentId,
            filedAt: new Date(entry.filedAt).toISOString()
          });
          // Upgrade cache: If we now have conversationId, store under that key too
          if (conversationId) {
            console.log("[findFiledEmailBySubject] Upgrading cache with conversationId:", conversationId.substring(0, 30) + "...");
            cache[conversationId] = entry;
            // Keep the subject-based entry for a while (don't delete)
            lsSet(JSON.stringify(cache));
            verification = lsGet();
            verifiedCache = verification ? JSON.parse(String(verification)) : {};
            upgradeSuccess = !!verifiedCache[conversationId];
            console.log("[findFiledEmailBySubject] Cache upgrade verification:", {
              success: upgradeSuccess,
              cacheSize: Object.keys(verifiedCache).length
            });
          }
          return [2 /*return*/, entry];
        }
        console.log("[findFiledEmailBySubject] ❌ No entry for this subject");
        return [2 /*return*/, null];
      } catch (e) {
        console.warn("[findFiledEmailBySubject] Failed:", e);
        return [2 /*return*/, null];
      }
      // removed by dead control flow

    });
  });
}
/**
 * Remove filed email from cache (e.g., if document was deleted)
 */
function removeFiledEmailFromCache(conversationId) {
  return __awaiter(this, void 0, void 0, function () {
    var raw, cache;
    return __generator(this, function (_a) {
      if (!conversationId) {
        return [2 /*return*/];
      }
      try {
        raw = lsGet();
        if (!raw) return [2 /*return*/];
        cache = JSON.parse(String(raw));
        delete cache[conversationId];
        lsSet(JSON.stringify(cache));
        console.log("[removeFiledEmailFromCache] Removed entry", {
          conversationId: conversationId.substring(0, 20) + "..."
        });
      } catch (e) {
        console.warn("[removeFiledEmailFromCache] Failed:", e);
      }
      return [2 /*return*/];
    });
  });
}

/***/ })

}]);
//# sourceMappingURL=src_utils_filedCache_ts.js.map