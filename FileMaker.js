(function () {

    /* Creating Singleton Object for File Maker Class */
    this.FileMaker = this.FileMaker || {};
    var fileMaker = this.FileMaker;

    fileMaker.Export = (function () {
        /* Private Members */
        var _isIE = (function () {
            var version = -1,
				appName = navigator.appName,
				userAgent = navigator.userAgent,
				regex = null;
            if (appName == 'Microsoft Internet Explorer') { // Upto IE10
                regex = new RegExp("MSIE ([0-9]{1,}[\.0-9]{0,})");
                if (regex.exec(userAgent) != null)
                    version = parseFloat(RegExp.$1);
            }
            else if (appName == 'Netscape') { // IE11
                regex = new RegExp("Trident/.*rv:([0-9]{1,}[\.0-9]{0,})");
                if (regex.exec(userAgent) != null)
                    version = parseFloat(RegExp.$1);

                if (version == -1) { // MS Edge
                    regex = new RegExp("Edge/([0-9]{1,}[\.0-9]{0,})");
                    if (regex.exec(userAgent) != null)
                        version = parseFloat(RegExp.$1);
                }
            }
            if (version > -1)
                return true;
            else
                return false;
        }()),
        _errors = {
            empty: {
                message: "{parameter} can't be null or empty."
            },
            elementTypeNotMatch: {
                message: "HTMLElement type doesn't matched, passed element type should be {type}."
            }
        },
        _defaults = {
            excelSheetName: 'Sheet 1'
        },

		/* Private Functions */

        _replacePlaceHolders = function (string, values) {
            var replacer = (function () {
                if (typeof values == "string")
                    return values;
                else if (typeof values == "object")
                    return function (match, parameter) { return values[parameter]; }
            }());
            return string.replace(/{(\w+)}/g, replacer);
        },
        _getElementFromDOM = function (element, elementType) {
            if (!element.nodeType) element = document.getElementById(element);

            if (element && String(element.nodeName).toLowerCase() == elementType.toLowerCase())
                return element;
            else {
                var exception = "Exception : " + _replacePlaceHolders(_errors.elementTypeNotMatch.message, String(elementType).toUpperCase());
                console.log(exception);
                alert(exception);
            }
        },
        _excelUsingIFrame = function (element, worksheetName) {
            var iframe = document.createElement("IFRAME");
            iframe.style.display = "none";
            document.body.appendChild(iframe);

            iframe.contentDocument.open("txt/html", "replace");
            iframe.contentDocument.write("<table>" + element.innerHTML + "</table>");
            iframe.contentDocument.close();
            iframe.focus();
            iframe.contentDocument.execCommand("SaveAs", true, worksheetName + '.xls');
        },
	    _excelUsingDataProtocol = function (element, worksheetName) {
	        var uri = 'data:application/vnd.ms-excel;base64,',
                template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>',
                base64 = function (string) { return window.btoa(unescape(encodeURIComponent(string))) };

	        var context = { worksheet: worksheetName || 'Worksheet', table: element.innerHTML }
	        window.location.href = uri + base64(_replacePlaceHolders(template, context))
	    };

        /* Public Functions */
        return {
            exportExcel: function (element, worksheetName) {
                if (!element) {
                    var exception = "Exception : " + _replacePlaceHolders(_errors.empty.message, "element");
                    console.log(exception);
                    alert(exception);
                }

                worksheetName = worksheetName || excelSheetName;
                element = _getElementFromDOM(element, "table");

                if (_isIE)
                    _excelUsingIFrame(element, worksheetName);
                else
                    _excelUsingDataProtocol(element, worksheetName);
            }
        };
    }());

    // Seal Object to stop more enhancements in object
    Object.seal(fileMaker);

}());