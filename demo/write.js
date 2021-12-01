const fs = require('fs');
const path = require('path');
const XLSX = require('../index.js').default;

/*
	表格样式
*/
let cellStyle = {
  font: {
    sz: 14,
    bold: true,
  },
  alignment: {
    horizontal: "center",
    vertical: "center",
  },
};

/*
	数据格式是一个二维数组，每一个数据代表一行
*/
let data = [
  [
    {
      v: '这是一个自定义表格',
      s: {
        font: {
          name: 'Microsoft YaHei',  // 字体
          sz: 15,                   // 字号
          color: {                  // 颜色
            rgb: 'FF0000',  
          },
          bold: true,               // 加粗
          underline: true,          // 下划线
          italic: true,             // 斜体
          strike: true,             // 删除线
          // outline:true,		      // 轮廓（无效）
          // shadow:true,			      // 阴影（无效）
          // vertAlign:true,		    // 垂直对齐方式(例如大字小字同时存在时，文字的对其基准)（无效）
        },
        alignment: {
          horizontal: "center",     // 水平居中
          vertical: "center",       // 垂直居中
          wrapText: true,           // 自动换行（默认不换行）
          // readingOrder:'2',	    // 阅读顺序（1、2没看出来有什么不同，貌似无效）
          // textRotation:30,		    // 文本旋转角度（没啥用的东东）
        },
        fill: {
          patternType: 'solid',     // 默认 solid，设为 none 时，fgColor 失效
          bgColor: {                // 无效
            rgb: "0000FF",
          },
          fgColor: {                // 背景色
            theme: "2",
            tint: "-0.25"
          }
        }
      }
    }
  ], 
  [
    {
      v: 1,
      s: cellStyle
    },
    {
      v: 2,
      s: cellStyle
    },
    {
      v: 3,
      s: cellStyle
    }
  ], 
  [
    {
      v: 4,
      s: cellStyle
    },
    {
      v: 5,
      s: cellStyle
    },
    {
      v: 6,
      s: cellStyle
    }
  ], 
  [
    {
      v: 7,
      s: cellStyle
    },
    {
      v: "百度一下",
      l: {
        Target: 'https://www.baidu.com',
        Tooltip: '这是一个 hover 提示',
      },
      s: {
        font: {
          color: {
            rgb: '0000FF',
          },
          underline: true,
        },
        alignment: cellStyle.alignment,
      }
    },
    {
      v: 36841,
      t: 'd',
      s: {
        ...cellStyle,
        numFmt: 'yyyy-mm-dd'
      }
    }
  ]
];

let options = {
  '!cols': [
    { wpx: 50 },
    { wpx: 50 },
    { wpx: 120 },
  ],
  '!rows': [
    { hpx: 30 },
    { hpx: 30 },
    { hpx: 40 },
    { hpx: 50 },
  ],
  '!merges': [
    {
      s: {
        c: 0,
        r: 0,
      },
      e: {
        c: 2,
        r: 0
      }
    }
  ],
};

let buffer = XLSX.build([
  {name: "Sheet111", data: data, options}
]);

fs.writeFileSync('123.xlsx', buffer);