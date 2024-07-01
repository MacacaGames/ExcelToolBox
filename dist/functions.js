/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	// The require scope
/******/ 	var __webpack_require__ = {};
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
/*!************************************!*\
  !*** ./src/functions/functions.ts ***!
  \************************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   add: function() { return /* binding */ add; },
/* harmony export */   clock: function() { return /* binding */ clock; },
/* harmony export */   currentTime: function() { return /* binding */ currentTime; },
/* harmony export */   increment: function() { return /* binding */ increment; },
/* harmony export */   logMessage: function() { return /* binding */ logMessage; },
/* harmony export */   tojson: function() { return /* binding */ tojson; }
/* harmony export */ });
/* global clearInterval, console, CustomFunctions, setInterval */

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

// /**
//  * convert data to Json
//  * @customfunction
//  * @param headers Json element name
//  * @param datas Json element value
//  * @returns {string} The Json of the data.
//  */
// export function tojson(headers: string, datas: string): string {
//   const objects = getJsonArrayFromData(headers, datas);
//   const json = JSON.stringify(objects);
//   if (json.length < 50000) {
//     return json;
//   }
//   return json;
// }

/**
 * convert data to Json
 * @customfunction
 * @param datas element value
 * @returns {string} The Json of the data.
 */
function tojson(datas) {
  const rows = datas.trim().split('\n');
  const array = rows.map(row => row.split(','));
  return dataArrayToJson(array);
}
function dataArrayToJson(data) {
  const headers = data[0];
  const rows = data.slice(1);
  const json = rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      let value = row[index];
      if (value === "NoData_Int" || value === "NoData_Enum") {
        value = 0;
      } else if (value === "NoData_Bool") {
        value = false;
      } else if (value === "NoData_Json") {
        value = {};
      } else if (value === "NoData_Array") {
        value = [];
      } else if (value === "NoData_Text") {
        value = "";
      } else if (!isNaN(value)) {
        // Convert numeric strings to numbers
        value = Number(value);
      }
      obj[header] = value;
    });
    return obj;
  });
  return JSON.stringify(json);
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);
  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);
  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
function logMessage(message) {
  console.log(message);
  return message;
}
CustomFunctions.associate("ADD", add);
CustomFunctions.associate("TOJSON", tojson);
CustomFunctions.associate("CLOCK", clock);
CustomFunctions.associate("INCREMENT", increment);
CustomFunctions.associate("LOG", logMessage);
/******/ })()
;
//# sourceMappingURL=functions.js.map