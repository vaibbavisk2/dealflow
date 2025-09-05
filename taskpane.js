Office.onReady(() => {
  const btn = document.getElementById("btn");
  const out = document.getElementById("out");
  btn.onclick = async () => {
    try {
      const item = Office.context.mailbox.item;
      // Works in read/compose where available
      const subject = item.subject && item.subject.getAsync
        ? await new Promise((res, rej) => item.subject.getAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(r.error)))
        : item.subject;
      out.textContent = "Subject: " + (subject || "(none)");
    } catch (e) {
      out.textContent = "Error: " + e.message;
    }
  };
});
