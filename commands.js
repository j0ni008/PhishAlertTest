Office.onReady(() => {});

function normalizeSubject(subject) {
  return (subject || "").trim();
}

function stripExistingPrefix(subject) {
  // Remove exactly these prefixes at the beginning
  return subject.replace(/^(::SPAM::|::PHISHING::|::Phishing::)\s*/i, "").trim();
}

function setSubjectPrefix(prefix) {
  const item = Office.context.mailbox.item;
  if (!item || !item.subject) return;

  item.subject.getAsync((res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) return;

    let subject = normalizeSubject(res.value);
    subject = stripExistingPrefix(subject);

    const newSubject = `${prefix} ${subject}`.trim();
    item.subject.setAsync(newSubject);
  });
}

// Must match manifest FunctionName
function markAsSpam(event) {
  setSubjectPrefix("::SPAM::");
  event.completed();
}

function markAsPhishing(event) {
  setSubjectPrefix("::Phishing::");
  event.completed();
}
