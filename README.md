# H5table2excel
* html export table to excel
 *  by Timemm 2021.5.8
 *  jQuery plugin to export an .xls file in browser from an HTML table
 *  Changes based on jQuery table2excel - v1.1.2
 *  compare with table2excel-v1.1.2
 *  support muti-sheet、muti-table
 *  support common css style (background-color、color、font-size、font-style .etc.)
 *  UTF-8
 *  https://github.com/Timemm/H5table2excel.git
 *  
 *  
 *  The original author infor
 *  jQuery table2excel - v1.1.2
 *  jQuery plugin to export an .xls file in browser from an HTML table
 *  https://github.com/rainabba/jquery-table2excel
 *  Made by rainabba
 *  Under MIT License



 # How to use

 * just add it in to html head
  * <script src="./js/jquery-3.3.1.min.js"></script>
  * <script src="./js/jquery.H5table2excel.js"></script>

 * option
  * $(selector).table2excel({
  *     exclude: ".noExl",//whats class tr has is not export to file
  *     filename: "hello_excel" + new Date().getTime() + ".xls",//whats name the file exported
  *     preserveColors: true ,// set to true if you want background colors and font colors preserved.The default is false
  *     preserveHtmlStyle: true //set to true if you want common html-css preserved(exclude color).The default is false
  * });

 * if you want more one table in one sheet or more one sheet
 * you can add attribute "data-SheetName" in table html

 # example
  *you can try : https://codepen.io/timemm/pen/xxqGEmG