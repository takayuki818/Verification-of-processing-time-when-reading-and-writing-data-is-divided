# Verification-of-processing-time-when-reading-and-writing-data-is-divided
10万行×100列のデータを行でn分割し配列格納＆セル範囲書き出しした場合の処理時間検証です  
1,2,4,5,8,10及び100,200,400,500,800,1000で分割してみましたが、  
（自身のPC環境下では）すべて20～22秒とほぼ変化がありませんでした。

※処理時間自体はApplicationの以下のプロパティ設定で高速化される余地があります。  
　'.Calculation = xlCalculationManual'  
　'.EnableEvents = False'  
　'.ScreenUpdating = False'  
