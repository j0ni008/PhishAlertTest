Office.onReady(() => {
  Office.actions.associate("reportPhish", reportPhish);
});

function reportPhish(event) {
  const item = Office.context.mailbox.item;

  item.forwardAsync(
    "security@barbaric.at",
    {
      comment: "Verdächtige E-Mail gemeldet"
    },
    function (result) {
      event.completed();
    }
  );
}
