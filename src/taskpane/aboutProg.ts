/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global document, Office, Word */

let list_of_changes = [
  "1) Улучшен механизм вставки ссылок на нумерацию",
  "2) Улучшен процесс отправки писем через мастер шаблонов (установлен таймаут в размере 5 сек на отправку одного письма - максимальное время ожидания ответа от сервера)",
  "3) Ускорен процесс тестирования полей в редакторе формул",
  "4) Исправлен родительный падеж женского рода в функции склонения ФИО в мастере шаблонов",
  "5) Добавлена защита от несканционированных изменений программного кода",
  "6) Добавлена интеграция Ассистента в текстовые файлы, для обеспечения работоспособности в случае невозможности его установки",
];

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("version").textContent = "Версия 2.73 от 12.06.2022";

    list_of_changes.forEach((change) => {
      let li = document.createElement("li");
      li.className = "ms-ListItem";

      let span = document.createElement("span");
      span.className = "ms-font-m";
      span.innerHTML = change;
      //<i class="ms-Icon ms-Icon--Ribbon ms-font-xl"></i>
      li.appendChild(span);
      document.getElementById("list-of-changes").appendChild(li);
    });
  }
});

// export async function run() {
//   return Word.run(async (context) => {
//     /**
//      * Insert your Word code here
//      */
//
//     // insert a paragraph at the end of the document.
//     const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
//
//     // change the paragraph color to blue.
//     paragraph.font.color = "blue";
//
//     await context.sync();
//   });
// }
