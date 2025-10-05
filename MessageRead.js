Office.initialize = function (reason) {
  $(document).ready(function () {
    var outlookItem = Office.context.mailbox.item;

    var parameters =
      "&messageId=" + encodeURIComponent(outlookItem.itemId) +
      "&subject=" + encodeURIComponent(outlookItem.subject) +
      "&from=" + encodeURIComponent(outlookItem.from.emailAddress) +
      "&fromname=" + encodeURIComponent(outlookItem.from.displayName) +
      "&receivedOn=" + encodeURIComponent(outlookItem.dateTimeCreated);

    document.getElementById('myApp').src += parameters;
  });
};
