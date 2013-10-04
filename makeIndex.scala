import scala.xml.XML
import scala.xml.Elem
import scala.io.Source

// 索引を付けるPPTファイルから抽出したテキストデータ
// <Slides>
//   <Slide>
//     <SlideNumber value="">
//     <SlideBody><!\CDATA[
//       スライドから抽出したテキストデータ
//       スライドノートから抽出したテキストデータ
//     ]]></SlideBody>
//   </Slide>
// </Slides>
val f = new java.io.File(args(0))
val xml = XML.loadFile(f)

// 索引用キーワードリスト
val s = Source.fromFile(args(1))
val keywords = try s.getLines.toList finally s.close // 単語のリストを読み込む

// XMLデータを読み込み、キーワードが出てきたページのページ番号リストを返す
def slideNumberList(k:String,xml:Elem):List[Int] = {
  val eachSlide:Seq[Int] = {
    for (s <- xml \ "Slide") yield {
      val target = s \ "SlideBody"
                             
      if (target.text.contains(k)) {
        (s \ "SlideNumber" \ "@value").text.toInt
      } else {
        0
      }
    }
  }
  eachSlide.toList.filter(n => n != 0)
}

// キーワードと番号リストのタプルを一覧で取得する
val index = for (k <- keywords) yield (k,slideNumberList(k,xml))

// 結果表示
for (i <- index) {
  val key = i._1
  val pageNumbers = i._2.mkString(",")
  println(s"${i._1}\t${pageNumbers}")
}
