Office.onReady(() => {
  document.getElementById("checkBtn").onclick = checkPairsInWord;
});

const pairs = [
  ['س','ث'], ['س','ذ'], ['س','ص'], ['س','ض'], ['س','ظ']
  // يمكنك إضافة المزيد لاحقًا
];

function checkPairsInWord() {
  Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();

    const text = range.text;
    if (!text) {
      showResult("يرجى تحديد نص داخل المستند.");
      return;
    }

    const words = text.split(/\s+/);
    let errors = [];

    // إزالة أي تمييز سابق
    range.clearAllFormatting();

    for (let word of words) {
      for (let pair of pairs) {
        const p = pair[0] + pair[1];
        if (word.includes(p)) {
          errors.push(`الكلمة «${word}» تحتوي على الثنائية (${p}) غير المقبولة.`);

          const searchResults = range.search(word, {matchCase: true, matchWholeWord: true});
          searchResults.load("items");
          await context.sync();

          searchResults.items.forEach(item => {
            item.font.highlightColor = "#FFD700"; // أصفر للتنبيه
          });

          break;
        }
      }
    }

    await context.sync();

    if (errors.length) {
      showResult(errors.join("<br>"), false);
    } else {
      showResult("النص المحدد سليم صوتيًا.", true);
    }
  });
}

function showResult(msg, valid = false) {
  const div = document.getElementById("result");
  div.innerHTML = msg;
  div.className = valid ? "result valid" : "result";
}
