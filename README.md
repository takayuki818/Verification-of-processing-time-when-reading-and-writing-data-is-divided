# Name
配列分割読み書き検証.xlsm
## Overview
10万行×100列のデータを行でn分割し配列格納＆セル範囲書き出しした場合の処理時間検証です。  
## Result
1,2,4,5,8,10及び100,200,400,500,800,1000で分割してみましたが、  
（自身のPC環境下では）すべて20～22秒とほぼ変化がありませんでした。
## Note
※処理時間自体は`Application`の以下のプロパティ変更で高速化される余地があります。  
　`.Calculation = xlCalculationManual`  
　`.EnableEvents = False`  
　`.ScreenUpdating = False`  
 ※上記プロパティ変更を実行する場合は、処理実行後に元に戻すことを忘れずに！
