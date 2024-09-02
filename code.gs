function buildAddOn(e) {
  var message = getMessageFromEvent(e);
  var remetente = message.getFrom();
  var remetenteNome = extractNameFromEmail(remetente);
  var remetenteEmail = extractEmailFromHeader(remetente);
  var remetenteFoto = getProfilePhoto(remetenteEmail);

  var card = CardService.newCardBuilder();

  // Cabeçalho com Nome e E-mail
  var headerSection = CardService.newCardSection();
  if (remetenteFoto) {
    headerSection.addWidget(CardService.newImage().setImageUrl(remetenteFoto).setAltText('Foto de ' + remetenteNome));
  }
  headerSection.addWidget(CardService.newTextParagraph().setText('<b>' + remetenteNome + '</b><br>' + remetenteEmail));
  
  card.addSection(headerSection);

  // Seção de Notas
  card.addSection(CardService.newCardSection()
    .setHeader("Notes")
    .addWidget(CardService.newTextParagraph().setText("No notes added.")));  // Placeholder para notas

  // Seção de Lembretes
  card.addSection(CardService.newCardSection()
    .setHeader("Reminders")
    .addWidget(CardService.newTextParagraph().setText("No reminders set.")));  // Placeholder para lembretes

  // Seção de Anexos
  var threads = GmailApp.search("from:" + remetenteEmail + " OR to:" + remetenteEmail);
  var sectionEmails = CardService.newCardSection().setHeader('Attachments');
  
  var totalAttachments = 0;
  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    messages.forEach(function(message) {
      var attachments = message.getAttachments();
      totalAttachments += attachments.length;
    });
  });

  sectionEmails.addWidget(CardService.newTextParagraph().setText(totalAttachments + " attachments"));
  card.addSection(sectionEmails);

  // Histórico de E-mails
  var emailSection = CardService.newCardSection().setHeader('Email History');
  var iconUrl = "https://www.flaticon.com/free-icons/mail";  // URL da pequena imagem que você quer adicionar ao lado dos emails
  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    messages.forEach(function(message) {
      var emailWidget = CardService.newKeyValue()
        .setIconUrl(iconUrl)  // Adiciona uma pequena imagem ao lado esquerdo do registro do email
        .setContent(message.getSubject())
        .setBottomLabel(message.getDate().toLocaleDateString() + ' ' + message.getDate().toLocaleTimeString())
        .setIcon(CardService.Icon.EMAIL);
      emailSection.addWidget(emailWidget);
    });
  });

  card.addSection(emailSection);

  return card.build();
}

function getMessageFromEvent(e) {
  var messageId = e.gmail.messageId;
  return GmailApp.getMessageById(messageId);
}

function extractNameFromEmail(email) {
  var name = email.match(/^(.*?)(?=\s*<)/);
  return name ? name[0] : email;
}

function extractEmailFromHeader(email) {
  var emailMatch = email.match(/<(.*)>/);
  return emailMatch ? emailMatch[1] : email;
}

function getProfilePhoto(email) {
  try {
    var url = "https://www.googleapis.com/admin/directory/v1/users/" + email + "/photos/thumbnail";
    var response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      }
    });
    var photo = JSON.parse(response.getContentText());
    return photo.url;
  } catch (e) {
    return null;  // Caso não consiga pegar a foto
  }
}
