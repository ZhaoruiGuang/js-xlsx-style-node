**js 处理 Excel 文件**

#### 本插件综合了 node-xlsx 和 xlsx-style，解决了两个插件之间的冲突，支持读写 xlsx 文件的同时，也支持设置简单的样式（目前仅支持 node 服务端的读写操作，不支持浏览器端，请按需选择）。

###### * 本文仅讲解了示例代码中用到的一些参数，更多参数可直接参考 xlsx-style 文档



#### 一、更新历史

​ **2021-12-01**  正式版第一版：1.1.0 发布。

​ **2021-11-29**  发现bug，处理中。。。

​ **2021-11-25**  修复几个源码bug，补充 readme 文档；

​ **2021-11-24**  第一次发布，测试；



#### 二、安装：

```js
npm install --save-dev js-xlsx-style-node
```



#### 三、写 xlsx 文件：

1. 先写一个简单数据表格：

   ```js
   const fs = require('fs');
   const path = require('path');
   
   /*
    本 demo 是在目录下相对路径引用的 XLSX ，使用者直接替换成 node_modules 引用路径即可:
    const XLSX = require('js-xlsx-style-node').default;
   */
   const = require('../index.js').default;
   
   /*
    数据格式是一个二维数组，每一个数据代表一行
   */
   let data = [
     [1,2,3],
     [4,5,6],
     [7,8,9]
   ]
   
   let buffer = XLSX.build([
     {name: "Sheet111", data: data}
   ]);
   
   fs.writeFileSync('123.xlsx',buffer);
   ```

   效果图：

   

   ![表格截图](https://p1.ssl.qhimg.com/t01eacce3b309843e56.png)

   一个简单的表格就这样诞生了！

   

2. 然后，再给表格增加点简单的样式，设置一下简单的列宽、行高、合并单元格等（通过 options 参数设置）：

   ```js
   const fs = require('fs');
   const path = require('path');
   const XLSX = require('../index.js').default;
   
   let data = [
     ['这是一个自定义表格'],
     [1,2,3],
     [4,5,6],
     [7,8,9]
   ];
   
   let options = {
     '!cols': [     
       { wpx: 50 }, 
    { wpx: 50 }, 
    { wpx: 60 }, 
     ],
     '!rows':[      
       { hpx: 30 },
    { hpx: 30 },
    { hpx: 40 },
    { hpx: 50 },
     ],
     '!merges':[
       {
      s:{
        c:0,
      r:0,
      },
      e:{
      c:2,
      r:0
      }
    }
     ],
   };
   
   let buffer = XLSX.build([
     {name: "Sheet111", data: data, options}
   ]);
   
   fs.writeFileSync('123.xlsx',buffer);
   ```

   效果图：

   

   ![表格截图](https://p1.ssl.qhimg.com/t018bfa5a33e0702810.png)

   参数都是啥意思呢？听我给你娓娓道来：

   - `'!cols'`  列宽，数组格式，数组每一项代表一列，列宽单位 wpx（像素）或 wch（厘米）；

   - `'!rows'`  行高，数组格式，数组每一项代表一行，行高单位 hpx（像素）或 hpt（磅）；

   - `'!merges'`  合并单元格，数组格式，每一项代表一个合并规则，s 代表合并的起始位置，e 代表合并的结束位置，c 代表列，r 代表行，示例中的合并单元格起始位置是：{c:0,r:0}，标识第一行第一列。结束位置是：{c:2,r:0}，代表第一行第三列。最终合并的效果就是那个表头了。

     

3. 接下来，再给表格增加更丰富样式，字体、字号、颜色、居中、背景色等：

   ```js
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
             // outline:true,       // 轮廓（无效）
             // shadow:true,      // 阴影（无效）
             // vertAlign:true,   // 垂直对齐方式(例如大字小字同时存在时，文字的对其基准)（无效）
           },
           alignment: {
             horizontal: "center",     // 水平居中
             vertical: "center",       // 垂直居中
             wrapText: true,           // 自动换行（默认不换行）
             // readingOrder:'2',     // 阅读顺序（1、2没看出来有什么不同，貌似无效）
             // textRotation:30,    // 文本旋转角度（没啥用的东东）
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
   ```

   效果图：

   
   
   ![表格截图](https://p3.ssl.qhimg.com/t01b6486bf180670444.png)
   
   

​ 具体参数含义就不一一解释了，看代码里的备注就行了，这里说几个注意事项：

- 给表格添加样式时，数据格式由简单的数组格式  `[1,2,3]`  变成了对象数组格式 `[ {v:1,s:cellStyle},{v:2,s:cellStyle},{v:3,s:cellStyle}]`，注意，每一个单元格项都变成这样的格式！各参数含义如下：
  - v：原始值；
  - t：内容类型，`s`表示string类型，`n`表示number类型，`b`表示boolean类型，`d`表示date类型，等等；
  - f：单元格公式，如`B2+B3`；
  - h：HTML内容；
  - w：格式化后的内容；
  - r：富文本内容 `rich text`；
  - l：单元格超链接对象，l.Target 表示链接，l.Tooltip 表示鼠标 hover 上去时的提示。*注意：Target 和 Tooltip  首字母大写！！官网 tooltip 首字母小写，但是亲测不生效！！！*
  - s：单元格样式/主题；
- 最后一个单元格，那是一个日期，值代表的是距离 1900年1月1日（*为什么是这一天，本人查来查去也没找到什么合理的解释，姑且猜测可能是因为这是 20 世纪开始的日子吧*🤣）的天数（*我们通常用的时间戳，是距离1970年1月1日的秒数或毫秒数*）。本插件读取 `xlsx` 文件时，日期读出的数据就是这样的格式。通过 `s` 参数下的 `numFmt` 属性可以将其格式化成正常的日期格式。

4. 另外，再介绍一些插件的其他方法：

   - 上面示例中，写文件时先调用 XLSX.build，再调用 fs.writeFileSync；也可以直接调用 XLSX.write 达到同样的效果：

     ```js
     let buffer = XLSX.build([
       {name: "Sheet111", data: data, options}
     ]);
     fs.writeFileSync('123.xlsx', buffer);
     ```

     等同于：

     ```js
     /* 同步写 */
     XLSX.write('123.xlsx',[{name: "Sheet111", data: data, options}]) 
     ```

     或者：

     ```js
     /* 异步写 */
     XLSX.writeAsync('123.xlsx',[{name: "Sheet111", data: data, options}],{/* 多个 sheet 共用的设置*/},()=>{
       console.log('这是异步写方法的回调~')
     }) 
     ```

   - 一个 xlsx 文件可能包含多个 sheet 表，同时写多个 sheet 只需在数组中多加几组数据即可：

     ```js
     let buffer = XLSX.build([
       {name: "Sheet1", data: data1, options1},
       {name: "Sheet2", data: data2, options2},
       {name: "Sheet3", data: data3, options3},
       {name: "Sheet4", data: data4, options4},
     ]);
     fs.writeFileSync('123.xlsx', buffer);
     ```

   - 本插件默认封装了几个方法，想调用原插件更多方法，可直接引用原插件的 XLSX 对象，但是建议不要用，因为可能会产生某种未知错误：

     ![截图](https://p5.ssl.qhimg.com/t014b1a95466cf0caf2.png)

     ​  

     ```js
     const XLSX = require('js-xlsx-style-node').default._XLSX;
     ```

   - 关于 options 里的参数，小编试了 `cellDates`、`tabSelected`、`showGridLines`、`bookSST`等参数，但是都没看出来有什么作用，可能是小编太笨了，就交给大家自己去研究吧。
   
5. “写” 方法小结

   ​  以上示例，涵盖了写 xlsx 文件的基本内容。我们日常写 xlsx 文件操作，无论你原本的数据是什么样的，只要按照示例中的格式将数据格式化，然后再调用插件方法，就可以生成一个带样式的表格了。有些属性，小编亲测无效；有些属性，小编测试完之后没发现什么变化，也不清楚有没有效果。更多更细致的用法，还请各位看官老爷们自己深挖吧~



#### 四、读 xlsx 文件：

​   源 xlsx 文件截图：



​   ![截图](https://p5.ssl.qhimg.com/t01b70b1b57f75faa95.png) 

​ 

1. 读成 json 格式：  

   ```js
   var fs = require('fs');
   var path = require('path');
   const XLSX = require('../index.js').default;
   
   let ws = XLSX.parse(path.resolve(__dirname,'read.xlsx'),{
    cellDates:true,     // 保留时间格式
   });
   console.log(ws[0].data)
   ```

   得到的数据：

   

   ![截图](https://p5.ssl.qhimg.com/t0105b0bdb52baffee4.png)

   *注：由于这个方法并未保留表格样式和超链接（实际上小编尝试了原插件里的各种方法，都没有保留样式和超链接的方式），所以我个人感觉对*

2. 读成 html 格式：

   ```js
   var fs = require('fs');
   var path = require('path');
   const XLSX = require('../index.js').default;
   
   let ws = XLSX.parseToHtml(path.resolve(__dirname,'read.xlsx'),{
    cellDates:true,     
    header:'',
    footer:'',
   });
   
   console.log(ws)
   
   ```

   可用通过参数 header、footer 来行添加 html 头尾：

   效果图：

   

   ![截图](https://p3.ssl.qhimg.com/t01f4181ff37b7534af.png)

   ![截图](https://p3.ssl.qhimg.com/t013bab80476cc83a07.png)

   - header 和 footer 为空时，得到的是一个 table 标签 html 片段；
   - 直接不配置 header 和 footer 参数时，得到的就是一个完整的 html 代码；
   - header 和 footer 配置值时，就以配置的字符串分别加载 table 片段的首尾；（这样配置自认为没啥实际用处）；

3. “读”小结

   ​  小编在读 xlsx 文件时，尝试了源插件里的各种方法，都没有得到带样式的数据。对于示例中的两种方法，在没有超链接的情况下，两者可以根据需要自行选择。但是有超链接的话，可能只能用读成 html 的方式了，因为 json 格式会导致超链接丢失。都不怎么好用，但是实在能力和时间有限，暂时也只能做到这样了😒。

   

##### 全文总结：

​ 本插件是在网上别的插件的基础上改的，改造之前，小编也试过直接用网上现有的一些插件，但都或多或少有点问题，满足不了需求：

1. `js-xlsx\node-xlsx`  不能设置样式；
2. `xlsx-style` 自己不能写文件，结合 `js-xlsx\node-xlsx` 还得处理源码冲突；
3. 好不容易找了一个能写样式的，还不支持写入超链接。。。
4. 等等

​ 最后，被逼无奈，只能硬着头皮看看源码，最终重要把样式和超链接的问题解决了，算是满足了一个表格的基本需求了。同时，还根据实际需要，给插件增加了几个读写方法，方便更多场景的使用。但是，由于时间和能力的问题，一些源文件里的方法和参数没有弄懂，所以也给大家遗留了很多困惑，对此，小编深感抱歉！美中不足，希望能给大家带来一点点帮助！



**GitHub 地址：**

​ https://github.com/ZhaoruiGuang/js-xlsx-style-node



人生艰难，撸码不易，您的支持是作者坚持下去的最大动力。

如果项目给您带来了一点便利或帮助，请留个star鼓励一下哟~~~❤️❤️❤️❤️❤️❤️

如有问题，可加微信15032361123（同为手机号）与作者交流。



