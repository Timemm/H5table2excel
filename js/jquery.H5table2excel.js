/*
 *  H5table2excel
 *  by Timemm 2021.5.8
 *  jQuery plugin to export an .xls file in browser from an HTML table
 *  Changes based on jQuery table2excel - v1.1.2
 *  compare with table2excel-v1.1.2
 *  support muti-sheet、muti-table
 *  support common css style (background-color、color、font-size、font-style .etc.)
 *  
 *  The original author infor
 *  jQuery table2excel - v1.1.2
 *  jQuery plugin to export an .xls file in browser from an HTML table
 *  https://github.com/rainabba/jquery-table2excel
 *  Made by rainabba
 *  Under MIT License
 */
(function ( $, window, document, undefined ) {
    var pluginName = "table2excel",

    defaults = {
        exclude: ".noExl",
        name: "Table2Excel",
        filename: "table2excel",
        fileext: ".xls",
        exclude_img: true,
        exclude_links: true,
        exclude_inputs: true,
        preserveColors: false,
        preserveHtmlStyle: false
    };

    // The actual plugin constructor
    function Plugin ( element, options ) {
        // console.log(element);
        this.element = element;
        this.settings = $.extend( {}, defaults, options );
        this._defaults = defaults;
        this._name = pluginName;
        this.init();
    }

    Plugin.prototype = {
        init: function () {
            var e = this;
            // console.log(e);

            var utf8Heading = "<meta http-equiv=\"content-type\" content=\"application/vnd.ms-excel; charset=UTF-8\">";

            e.tableRows = [];

			// Styling variables
			var additionalStyles = "";
			var compStyle = null;

            // get contents of table except for exclude
            $(e.element).each( function(i,o) {
                var tempRows = "";
                $(o).find("tr").not(e.settings.exclude).each(function (i,p) {
					// Reset for this row
					additionalStyles = "";
					
					// Preserve background and text colors on the row
					if(e.settings.preserveColors){
						compStyle = getComputedStyle(p);
						additionalStyles += (compStyle && compStyle.backgroundColor ? "background-color: " + compStyle.backgroundColor + ";" : "");
                        additionalStyles += (compStyle && compStyle.color ? "color: " + compStyle.color + ";" : "");
					}
                    if (e.settings.preserveHtmlStyle) {
                        //modify by wzy  to support other common style.
                        compStyle = getComputedStyle(p);
                        additionalStyles += (compStyle && compStyle.fontSize ? "font-size: " + compStyle.fontSize + ";" : "");
                        additionalStyles += (compStyle && compStyle.textAlign ? "text-align: " + compStyle.textAlign + ";" : "");
                        additionalStyles += (compStyle && compStyle.fontWeight ? "font-weight: " + compStyle.fontWeight + ";" : "");
                        additionalStyles += (compStyle && compStyle.fontStyle ? "font-style: " + compStyle.fontStyle + ";" : "");
                        additionalStyles += (compStyle && compStyle.width ? "width: " + compStyle.width + ";" : "auto;");
                        additionalStyles += (compStyle && compStyle.height ? "height: " + compStyle.height + ";" : "auto;");
                        additionalStyles += (compStyle && compStyle.border ? "border: " + compStyle.border + ";" : "");
                    }

					// Create HTML for Row
                    tempRows += "<tr style='" + additionalStyles + "'>";
                    
                    // Loop through each TH and TD
                    $(p).find("td,th").not(e.settings.exclude).each(function (i,q) { // p did not exist, I corrected
						
						// Reset for this column
						additionalStyles = "";
						
						// Preserve background and text colors on the row
						if(e.settings.preserveColors){
							compStyle = getComputedStyle(q);
							additionalStyles += (compStyle && compStyle.backgroundColor ? "background-color: " + compStyle.backgroundColor + ";" : "");
							additionalStyles += (compStyle && compStyle.color ? "color: " + compStyle.color + ";" : "");
						}
                        if (e.settings.preserveHtmlStyle) {
                            //modify by wzy  to support other common style.
                            compStyle = getComputedStyle(q);
                            additionalStyles += (compStyle && compStyle.fontSize ? "font-size: " + compStyle.fontSize + ";" : "");
                            additionalStyles += (compStyle && compStyle.textAlign ? "text-align: " + compStyle.textAlign + ";" : "");
                            additionalStyles += (compStyle && compStyle.fontWeight ? "font-weight: " + compStyle.fontWeight + ";" : "");
                            additionalStyles += (compStyle && compStyle.fontStyle ? "font-style: " + compStyle.fontStyle + ";" : "");
                            additionalStyles += (compStyle && compStyle.width ? "width: " + compStyle.width + ";" : "auto;");
                            additionalStyles += (compStyle && compStyle.height ? "height: " + compStyle.height + ";" : "auto;");
                            additionalStyles += (compStyle && compStyle.border ? "border: " + compStyle.border + ";" : "");
                        }

                        var rc = {
                            rows: $(this).attr("rowspan"),
                            cols: $(this).attr("colspan"),
                            flag: $(q).find(e.settings.exclude)
                        };

                        if( rc.flag.length > 0 ) {
                            tempRows += "<td> </td>"; // exclude it!!
                        } else {
                            tempRows += "<td";
                            if( rc.rows > 0) {
                                tempRows += " rowspan='" + rc.rows + "' ";
                            }
                            if( rc.cols > 0) {
                                tempRows += " colspan='" + rc.cols + "' ";
                            }
                            if(additionalStyles){
								tempRows += " style='" + additionalStyles + "'";
							}
                            tempRows += ">" + $(q).html() + "</td>";
                        }
                    });

                    tempRows += "</tr>";

                });
                // exclude img tags
                if(e.settings.exclude_img) {
                    tempRows = exclude_img(tempRows);
                }

                // exclude link tags
                if(e.settings.exclude_links) {
                    tempRows = exclude_links(tempRows);
                }

                // exclude input tags
                if(e.settings.exclude_inputs) {
                    tempRows = exclude_inputs(tempRows);
                }
                e.tableRows.push(tempRows);
            });
            // console.log(e.settings.sheetName);

            e.tableToExcel(e.tableRows, e.settings.name, e.settings.sheetName);
        },

        tableToExcel: function (table, name, sheetName) {
            var e = this, fullTemplate="", i, link, a;
            var html_start = `<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">`
    , template_ExcelWorksheet = `<x:ExcelWorksheet><x:Name>{SheetName}</x:Name><x:WorksheetSource HRef="sheet{SheetIndex}.htm"/><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>`
    , template_ListWorksheet = `<o:File HRef="sheet{SheetIndex}.htm"/>`
    , template_HTMLWorksheet = `
------=_NextPart_dummy
Content-Location: sheet{SheetIndex}.htm
Content-Type: text/html; charset=UTF-8

` + html_start + `
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <link id="Main-File" rel="Main-File" href="../WorkBook.htm">
    <link rel="File-List" href="filelist.xml">
</head>
<body>{SheetContent}</body>
</html>`
    , template_WorkBook = `MIME-Version: 1.0
X-Document-Type: Workbook
Content-Type: multipart/related; boundary="----=_NextPart_dummy"

------=_NextPart_dummy
Content-Location: WorkBook.htm
Content-Type: text/html; charset=UTF-8

` + html_start + `
<head>
<meta name="Excel Workbook Frameset">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<link rel="File-List" href="filelist.xml">
<!--[if gte mso 9]><xml>
 <x:ExcelWorkbook>
    <x:ExcelWorksheets>{ExcelWorksheets}</x:ExcelWorksheets>
    <x:ActiveSheet>0</x:ActiveSheet>
 </x:ExcelWorkbook>
</xml><![endif]-->
</head>
<frameset>
    <frame src="sheet0.htm" name="frSheet">
    <noframes><body><p>This page uses frames, but your browser does not support them.</p></body></noframes>
</frameset>
</html>
{HTMLWorksheets}
Content-Location: filelist.xml
Content-Type: text/xml; charset="utf-8"

<xml xmlns:o="urn:schemas-microsoft-com:office:office">
    <o:MainFile HRef="../WorkBook.htm"/>
    {ListWorksheets}
    <o:File HRef="filelist.xml"/>
</xml>
------=_NextPart_dummy--
`
            e.format = function (s, c) {
                return s.replace(/{(\w+)}/g, function (m, p) {
                    return c[p];
                });
            };

            sheetName = typeof sheetName === "undefined" ? "Sheet" : sheetName;

            e.ctx = {
                worksheet: name || "Worksheet",
                table: table,
                sheetName: sheetName
            };

            // console.log(e.element);
            // console.log(e.ctx);
            // console.log(fullTemplate);

            //modify by wzy 
            var context_sheet = {};
            var context_WorkBook = {
                ExcelWorksheets:'',
                HTMLWorksheets: '',
                ListWorksheets: ''
            };
            if ( $.isArray(table) ) {
                 Object.keys(table).forEach(function(i){
                    let SheetName = $($(e.element)[i]).attr('data-SheetName');
                    // console.log($.trim(SheetName));

                    if ($.trim(SheetName).replace(/\s/g,"") === '') {//无data-SheetName不导出

                    }else{
                        if (typeof(context_sheet[SheetName.replace(/\s/g,"")]) == "undefined") {//新建sheet
                            context_WorkBook.ExcelWorksheets += e.format(template_ExcelWorksheet, {
                                SheetIndex: i,
                                SheetName: SheetName
                            });
                            context_WorkBook.HTMLWorksheets += e.format(template_HTMLWorksheet, {
                                SheetIndex: i,
                                SheetContent: '{'+SheetName.replace(/\s/g,"")+'}'
                            });
                            context_WorkBook.ListWorksheets += e.format(template_ListWorksheet, {
                                SheetIndex: i
                            });
                            context_sheet[SheetName.replace(/\s/g,"")] = "<table>" + "{table" + i + "}" + "</table>";
                            
                        }else{//已有该sheet
                            context_sheet[SheetName.replace(/\s/g,"")] += "<div></div><table>" + "{table" + i + "}" + "</table>";

                        }
                        
                    }
                });
            }
            fullTemplate =  e.format(template_WorkBook, context_WorkBook);
            // console.log(fullTemplate);
            fullTemplate =  e.format(fullTemplate, context_sheet);
            //modify end
            // console.log(fullTemplate,context_sheet);
            for (i in table) {
                // console.log(table[i])
                e.ctx["table" + i] = table[i];
            }
            delete e.ctx.table;

            var isIE = navigator.appVersion.indexOf("MSIE 10") !== -1 || (navigator.userAgent.indexOf("Trident") !== -1 && navigator.userAgent.indexOf("rv:11") !== -1); // this works with IE10 and IE11 both :)
            //if (typeof msie !== "undefined" && msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./))      // this works ONLY with IE 11!!!
            if (isIE) {
                if (typeof Blob !== "undefined") {
                    //use blobs if we can
                    fullTemplate = e.format(fullTemplate, e.ctx); // with this, works with IE
                    fullTemplate = [fullTemplate];
                    //convert to array
                    var blob1 = new Blob(fullTemplate, { type: "text/html" });
                    window.navigator.msSaveBlob(blob1, getFileName(e.settings) );
                } else {
                    //otherwise use the iframe and save
                    //requires a blank iframe on page called txtArea1
                    txtArea1.document.open("text/html", "replace");
                    txtArea1.document.write(e.format(fullTemplate, e.ctx));
                    txtArea1.document.close();
                    txtArea1.focus();
                    sa = txtArea1.document.execCommand("SaveAs", true, getFileName(e.settings) );
                }

            } else {
                var blob = new Blob([e.format(fullTemplate, e.ctx)], {type: "application/vnd.ms-excel"});
                window.URL = window.URL || window.webkitURL;
                link = window.URL.createObjectURL(blob);
                a = document.createElement("a");
                a.download = getFileName(e.settings);
                a.href = link;
                // a.target = '_blank';
                document.body.appendChild(a);

                a.click();

                document.body.removeChild(a);
            }

            return true;
        }
    };

    function getFileName(settings) {
        return ( settings.filename ? settings.filename : "table2excel" );
    }

    // Removes all img tags
    function exclude_img(string) {
        var _patt = /(\s+alt\s*=\s*"([^"]*)"|\s+alt\s*=\s*'([^']*)')/i;
        return string.replace(/<img[^>]*>/gi, function myFunction(x){
            var res = _patt.exec(x);
            if (res !== null && res.length >=2) {
                return res[2];
            } else {
                return "";
            }
        });
    }

    // Removes all link tags
    function exclude_links(string) {
        return string.replace(/<a[^>]*>|<\/a>/gi, "");
    }

    // Removes input params
    function exclude_inputs(string) {
        var _patt = /(\s+value\s*=\s*"([^"]*)"|\s+value\s*=\s*'([^']*)')/i;
        return string.replace(/<input[^>]*>|<\/input>/gi, function myFunction(x){
            var res = _patt.exec(x);
            if (res !== null && res.length >=2) {
                return res[2];
            } else {
                return "";
            }
        });
    }

    $.fn[ pluginName ] = function ( options ) {
        var e = this;

        e.each(function(index) {
            // console.log(e,index);
            if ( !$.data( e, "plugin_" + pluginName ) ) {
                // console.log(e);

                //modify by wzy 2021.5.7
                //change this to e 
                //to support muti-sheet in one xls.

                // $.data( e, "plugin_" + pluginName, new Plugin( this, options ) );
                $.data( e, "plugin_" + pluginName, new Plugin( e, options ) );
                //modify end
            }
        });

        // chain jQuery functions
        // console.log(e,$.data( e, "plugin_" + pluginName ).element);
        return e;
    };

})( jQuery, window, document );
