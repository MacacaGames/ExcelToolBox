/******/ (function() { // webpackBootstrap
/******/ 	var __webpack_modules__ = ({

/***/ "./node_modules/html-loader/dist/runtime/getUrl.js":
/*!*********************************************************!*\
  !*** ./node_modules/html-loader/dist/runtime/getUrl.js ***!
  \*********************************************************/
/***/ (function(module) {

"use strict";


module.exports = function (url, options) {
  if (!options) {
    // eslint-disable-next-line no-param-reassign
    options = {};
  }
  if (!url) {
    return url;
  }

  // eslint-disable-next-line no-underscore-dangle, no-param-reassign
  url = String(url.__esModule ? url.default : url);
  if (options.hash) {
    // eslint-disable-next-line no-param-reassign
    url += options.hash;
  }
  if (options.maybeNeedQuotes && /[\t\n\f\r "'=<>`]/.test(url)) {
    return "\"".concat(url, "\"");
  }
  return url;
};

/***/ }),

/***/ "./assets/logo-filled.png":
/*!********************************!*\
  !*** ./assets/logo-filled.png ***!
  \********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

"use strict";
module.exports = __webpack_require__.p + "assets/logo-filled.png";

/***/ }),

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

"use strict";
module.exports = __webpack_require__.p + "1fda685b81e1123773f6.css";

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
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = __webpack_modules__;
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	!function() {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = function(module) {
/******/ 			var getter = module && module.__esModule ?
/******/ 				function() { return module['default']; } :
/******/ 				function() { return module; };
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	}();
/******/ 	
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
/******/ 	/* webpack/runtime/publicPath */
/******/ 	!function() {
/******/ 		var scriptUrl;
/******/ 		if (__webpack_require__.g.importScripts) scriptUrl = __webpack_require__.g.location + "";
/******/ 		var document = __webpack_require__.g.document;
/******/ 		if (!scriptUrl && document) {
/******/ 			if (document.currentScript)
/******/ 				scriptUrl = document.currentScript.src;
/******/ 			if (!scriptUrl) {
/******/ 				var scripts = document.getElementsByTagName("script");
/******/ 				if(scripts.length) {
/******/ 					var i = scripts.length - 1;
/******/ 					while (i > -1 && (!scriptUrl || !/^http(s?):/.test(scriptUrl))) scriptUrl = scripts[i--].src;
/******/ 				}
/******/ 			}
/******/ 		}
/******/ 		// When supporting browsers where an automatic publicPath is not supported you must specify an output.publicPath manually via configuration
/******/ 		// or pass an empty string ("") and set the __webpack_public_path__ variable from your code to use your own logic.
/******/ 		if (!scriptUrl) throw new Error("Automatic publicPath is not supported in this browser");
/******/ 		scriptUrl = scriptUrl.replace(/#.*$/, "").replace(/\?.*$/, "").replace(/\/[^\/]+$/, "/");
/******/ 		__webpack_require__.p = scriptUrl;
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/jsonp chunk loading */
/******/ 	!function() {
/******/ 		__webpack_require__.b = document.baseURI || self.location.href;
/******/ 		
/******/ 		// object to store loaded and loading chunks
/******/ 		// undefined = chunk not loaded, null = chunk preloaded/prefetched
/******/ 		// [resolve, reject, Promise] = chunk loading, 0 = chunk loaded
/******/ 		var installedChunks = {
/******/ 			"taskpane": 0
/******/ 		};
/******/ 		
/******/ 		// no chunk on demand loading
/******/ 		
/******/ 		// no prefetching
/******/ 		
/******/ 		// no preloaded
/******/ 		
/******/ 		// no HMR
/******/ 		
/******/ 		// no HMR manifest
/******/ 		
/******/ 		// no on chunks loaded
/******/ 		
/******/ 		// no jsonp function
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry need to be wrapped in an IIFE because it need to be isolated against other entry modules.
!function() {
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.ts ***!
  \**********************************/
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("selectRangeToJsonButton").onclick = selectRangeToJson;
  document.getElementById("selectWholeSheetToJsonButton").onclick = selectWholeSheetToJson;
  document.getElementById("insertNoDataIntButton").onclick = insertNoDataInt;
  document.getElementById("insertNoDataEnumButton").onclick = insertNoDataEnum;
  document.getElementById("insertNoDataBoolButton").onclick = insertNoDataBool;
  document.getElementById("insertNoDataArrayButton").onclick = insertNoDataArray;
  document.getElementById("insertNoDataTextButton").onclick = insertNoDataText;
});
async function selectRangeToJson() {
  try {
    const rangeInput = document.getElementById("rangeInput").value.trim();
    if (!rangeInput) {
      console.error('Please enter a valid range.');
      return;
    }
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const excelRange = sheet.getRange(rangeInput);
      excelRange.load("values");
      await context.sync();
      const data = excelRange.values;
      await dataArrayToJson(data);
    });
  } catch (error) {
    console.error(error);
  }
}
async function selectWholeSheetToJson() {
  try {
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const excelRange = sheet.getUsedRange();
      excelRange.load("values");
      await context.sync();
      const data = excelRange.values;
      await dataArrayToJson(data);
    });
  } catch (error) {
    console.error(error);
  }
}
async function dataArrayToJson(data) {
  const headers = data[0];
  const rows = data.slice(1);
  let jsonArray = [];
  for (let i = 0; i < rows.length; i += 50) {
    const batch = rows.slice(i, i + 50);
    const jsonBatch = batch.map(row => {
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
        } else {
          try {
            const parsedValue = JSON.parse(value);
            if (typeof parsedValue === 'object') {
              value = parsedValue;
            }
          } catch (e) {}
        }
        obj[header] = value;
      });
      return obj;
    });
    jsonArray = jsonArray.concat(jsonBatch);

    // Intermediate sync to keep the UI responsive
    await new Promise(resolve => setTimeout(resolve, 50));
  }
  const jsonString = JSON.stringify(jsonArray, null, 2);
  await openJsonDialog(jsonString);
}
async function openJsonDialog(jsonString) {
  localStorage.setItem('json', jsonString);
  Office.context.ui.displayDialogAsync(location.origin + '/taskpane/dialog.html', {
    height: 50,
    width: 50
  }, result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error('Failed to open dialog:', result.error.message);
    }
  });
}
async function insertNoDataInt() {
  await insertData("NoData_Int");
}
async function insertNoDataEnum() {
  await insertData("NoData_Enum");
}
async function insertNoDataBool() {
  await insertData("NoData_Bool");
}
async function insertNoDataJson() {
  await insertData("NoData_Json");
}
async function insertNoDataArray() {
  await insertData("NoData_Array");
}
async function insertNoDataText() {
  await insertData("NoData_Text");
}
async function insertData(data) {
  try {
    await Excel.run(async context => {
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.values = data;
      await context.sync();
    });
  } catch (error) {
    console.error("Error:", error);
  }
}
}();
// This entry need to be wrapped in an IIFE because it need to be in strict mode.
!function() {
"use strict";
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../node_modules/html-loader/dist/runtime/getUrl.js */ "./node_modules/html-loader/dist/runtime/getUrl.js");
/* harmony import */ var _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0__);
// Imports

var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.css */ "./src/taskpane/taskpane.css"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_1___ = new URL(/* asset import */ __webpack_require__(/*! ../../assets/logo-filled.png */ "./assets/logo-filled.png"), __webpack_require__.b);
// Module
var ___HTML_LOADER_REPLACEMENT_0___ = _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default()(___HTML_LOADER_IMPORT_0___);
var ___HTML_LOADER_REPLACEMENT_1___ = _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default()(___HTML_LOADER_IMPORT_1___);
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\n<!-- This file shows how to design a first-run page that provides a welcome screen to the user about the features of the add-in. -->\n\n<!DOCTYPE html>\n<html>\n\n<head>\n    <meta charset=\"UTF-8\" />\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n    <title>Contoso Task Pane Add-in</title>\n\n    <!-- Office JavaScript API -->\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js\"><" + "/script>\n\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\n    <link rel=\"stylesheet\" href=\"https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/11.0.0/css/fabric.min.css\"/>\n\n    <!-- Template styles -->\n    <link href=\"" + ___HTML_LOADER_REPLACEMENT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\n</head>\n\n<body class=\"ms-font-m ms-welcome ms-Fabric\">\n    <header class=\"ms-welcome__header ms-bgColor-neutralLighter\">\n        <img width=\"90\" height=\"90\" src=\"" + ___HTML_LOADER_REPLACEMENT_1___ + "\" alt=\"Contoso\" title=\"Contoso\" />\n        <h1 class=\"ms-font-su\">Macaca DataSheet</h1>\n    </header>\n    <section id=\"sideload-msg\" class=\"ms-welcome__main\">\n        <h2 class=\"ms-font-xl\">Please <a target=\"_blank\" href=\"https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing\">sideload</a> your add-in to see app body.</h2>\n    </section>\n    <main id=\"app-body\" class=\"ms-welcome__main\" style=\"display: none;\">\n        <h2 class=\"ms-font-xl\"> Discover what Office Add-ins can do for you today! </h2>\n        <ul class=\"ms-List ms-welcome__features\">\n            <li class=\"ms-ListItem\">\n                <i class=\"ms-Icon ms-Icon--Ribbon ms-font-xl\"></i>\n                <span class=\"ms-font-m\">Achieve more with Office integration</span>\n            </li>\n            <li class=\"ms-ListItem\">\n                <i class=\"ms-Icon ms-Icon--Unlock ms-font-xl\"></i>\n                <span class=\"ms-font-m\">Unlock features and functionality</span>\n            </li>\n            <li class=\"ms-ListItem\">\n                <i class=\"ms-Icon ms-Icon--Design ms-font-xl\"></i>\n                <span class=\"ms-font-m\">Create and visualize like a pro</span>\n            </li>\n        </ul>\n        <p class=\"ms-font-l\">Modify the source files, then click <b>Run</b>.</p>\n        <div role=\"button\" id=\"run\" class=\"ms-welcome__action ms-Button ms-Button--hero ms-font-xl\">\n            <span class=\"ms-Button-label\">Run</span>\n        </div>\n        <div>\n            <label for=\"rangeInput\">Enter range (e.g., A1:C3):</label>\n            <input type=\"text\" id=\"rangeInput\" />\n            <button id=\"selectRangeToJsonButton\">Get Json</button>\n        </div>\n        <div>\n            <label for=\"rangeInput\">Whole sheet to Json</label>\n            <button id=\"selectWholeSheetToJsonButton\">Get Json</button>\n        </div>\n        <div>\n            <label>Insert no data</label>\n            <button id=\"insertNoDataIntButton\">Int</button>\n            <button id=\"insertNoDataEnumButton\">Enum</button>\n            <button id=\"insertNoDataBoolButton\">Bool</button>\n            <button id=\"insertNoDataArrayButton\">Array</button>\n            <button id=\"insertNoDataTextButton\">Text</button>\n        </div>\n        <p><label id=\"item-subject\"></label></p>\n        <pre id=\"jsonContent\"></pre>\n    </main>\n\n</body>\n\n</html>\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=taskpane.js.map