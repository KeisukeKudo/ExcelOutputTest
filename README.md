# ExcelOutputTest
とある理由でExcel出力のパフォーマンスをテスト  
使ったライブラリは[EPPlus](http://epplus.codeplex.com/ "EPPlus")  

# 条件
200列  
30,001行(1行目は列名)  
6,000,200セル    

# 結果
作成時間: 209,339ms(3分半くらい)  
![result1](https://github.com/KeisukeKudo/ImageStorage/blob/master/ExcelOutputTestResult1.png)  

作成Excelファイル  
![result2](https://github.com/KeisukeKudo/ImageStorage/blob/master/ExcelOutputTestResult2.png)  
