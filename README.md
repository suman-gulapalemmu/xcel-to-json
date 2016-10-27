# xcel-to-json
Node module which can convert an excel file to JSON

There are quite a few npm modules which will do this job : [excel-as-json](https://www.npmjs.com/package/excel-as-json) , [xls-to-json](https://www.npmjs.com/package/xls-to-json) 

[excel-as-json](https://www.npmjs.com/package/excel-as-json) is completely written in coffeescript and it just shows nightmares when trying to install that node module in windows. You get all sorts of errors / warnings to install like `node-gyp rebuild` issues
But this module offers great features like creating nested objects etc which other modules lack, only if installation was not so cumbersome this module would have been best one to use.
Neither do i have knowledge of coffeescript to dig in and change the source code nor i found enough support online to rectify `node-gyp rebuild` issue.

xls-to-json is simple and i have used this as reference to build my own module which can address all features provided by [excel-as-json](https://www.npmjs.com/package/excel-as-json)
