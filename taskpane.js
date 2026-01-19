Office.onReady(function() {
  var btn = document.getElementById("checkBtn");
  if (btn) {
    btn.onclick = checkPairsInWord;
  }
});

var pairs = [
  ['س', 'ث'], ['س', 'ذ'], ['س', 'ص'], ['س', 'ض'], ['س', 'ظ']
];

function checkPairsInWord() {
  Word.run(function (context) {
    var range = context.document.getSelection();
    range.load("text");

    // الخطوة 1: جلب النص
    return context.sync().then(function () {
      var text = range.text;
      
      if (!text) {
        showResult("يرجى تحديد نص داخل المستند.");
        return;
      }

      // تنظيف التنسيق القديم (إزالة الألوان السابقة)
      range.clearAllFormatting();

      var words = text.split(/\s+/);
      var errors = [];
      var searchResultsArray = []; // لتخزين نتائج البحث للمعالجة لاحقاً

      // الخطوة 2: تحليل الكلمات وتجهيز البحث
      for (var i = 0; i < words.length; i++) {
        var word = words[i];
        
        for (var j = 0; j < pairs.length; j++) {
          var p = pairs[j][0] + pairs[j][1];
          
          // بديل لـ .includes()
          if (word.indexOf(p) !== -1) {
             errors.push('الكلمة «' + word + '» تحتوي على الثنائية (' + p + ') غير المقبولة.');
             
             // البحث عن الكلمة داخل الوورد لتلوينها
             // (matchWholeWord: true لضمان تلوين الكلمة بالكامل فقط)
             var searchResult = range.search(word, { matchCase: true, matchWholeWord: true });
             searchResult.load("items");
             searchResultsArray.push(searchResult);
             
             break; // الانتقال للكلمة التالية
          }
        }
      }

      // إذا لم تكن هناك أخطاء
      if (errors.length === 0) {
         showResult("النص المحدد سليم صوتيًا.", true);
         return; 
      }

      // الخطوة 3: تنفيذ التلوين (Sync واحد للجميع لزيادة السرعة)
      return context.sync().then(function() {
        // الآن وقد تم تحميل نتائج البحث، نقوم بتلوينها
        for (var k = 0; k < searchResultsArray.length; k++) {
           var items = searchResultsArray[k].items;
           for (var m = 0; m < items.length; m++) {
             items[m].font.highlightColor = "#FFD700"; // أصفر
           }
        }
        
        // عرض الرسائل للمستخدم
        showResult(errors.join("<br>"), false);
      });
    });
  })
  .catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function showResult(msg, valid) {
  var div = document.getElementById("result");
  div.innerHTML = msg;
  div.className = valid ? "result valid" : "result";
}
