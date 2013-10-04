@rem PowerPointファイルから、索引情報を作成する
@rem 使用法
@rem PPT2Index.bat PowerPointファイル テキストデータ抽出結果ファイル 索引用キーワードファイル
@rem ex: > PPT2Index.bat CourseWare.pptx output.xml keywords.txt

cscript PPT2Text.vbs %1 //nologo > %2 
scala makeIndex.scala %2 %3