// // /*
// // * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// // * See LICENSE in the project root for license information.

// * This file specifically used for Outlook on Desktop.
// */

function onAppointmentSendHandler(event) {
  Office.context.mailbox.item.categories.getAsync(async function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const categories = asyncResult.value;

      if (categories && categories.length > 0) {
        if (
          categories.some((r) => r.displayName.includes("Nasdaq CMS Meetings"))
        ) {
          //console.log("Category", categories);

          const mailboxItem = Office.context.mailbox.item;

          let StartTime = "";
          let EndTime = "";
          let Timezone = "";
          let Location = "";
          let AttachmentCount = "";
          let AttachedFiles = "";
          let Subject = "";

          mailboxItem.getAttachmentsAsync(function (result) {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              console.error(result.error.message);
            } else {
              AttachmentCount = result.value.length.toString();
              if (result.value.length > 0) {
                let fileNameString = "";
                result.value.forEach((file) => {
                  fileNameString = fileNameString + file.name;
                });
                AttachedFiles = fileNameString;
              }
            }
          });

          mailboxItem.location.getAsync((result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              console.error(`Failed to get Location ${result.error.message}`);
              return;
            } else {
              Location = result.value;
            }
          });

          mailboxItem.start.getAsync(async (asyncResult) => {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
              console.error(
                "Failed to get start time ",
                asyncResult.error.message
              );
            } else {
              let startDate = asyncResult.value.toString();
              let startdateTime = getFormattedDate(startDate);
              StartTime = startdateTime.formattedDate;
              Timezone = startdateTime.timezoneID;
            }
          });

          mailboxItem.end.getAsync(async (asyncResult) => {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
              console.error(
                "Failed to get end time ",
                asyncResult.error.message
              );
            } else {
              let endDate = asyncResult.value.toString();
              let enddateTime = getFormattedDate(endDate);
              EndTime = enddateTime.formattedDate;
            }
          });

          mailboxItem.subject.getAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.error("Failed to get subject", asyncResult.error.message);
            } else {
              Subject = asyncResult.value;
            }
          });

          mailboxItem.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            const synched = props.get("synched");
            const subject = props.get("subject");
            const deleted = props.get("deleted");

            const startTime = props.get("startTime");
            const endTime = props.get("endTime");
            const timezone = props.get("timezone");
            const location = props.get("location");
            const attachmentCount = props.get("attachmentCount");
            const attachedFiles = props.get("attachedFiles");

            // console.log("startTime", startTime);
            // console.log("endTime", endTime);
            // console.log("timezone", timezone);
            // console.log("location", location);
            // console.log("attachmentCount", attachmentCount);
            // console.log("attachedFiles", attachedFiles);
            // console.log("subject", attachedFiles);

            // console.log("********************")
            // console.log("StartTime", StartTime);
            // console.log("EndTime", EndTime);
            // console.log("Timezone", Timezone);
            // console.log("Location", Location);
            // console.log("AttachmentCount", AttachmentCount);
            // console.log("AttachedFiles", AttachedFiles);
            // console.log("Subject", subject);

            if (deleted === "true") {
              event.completed({ allowEvent: true });
            }

            if (Subject === "") {
              event.completed({
                allowEvent: false,
                errorMessage: "Title cannot be blank. Please add ",
              });
            } else if (synched === "true" || deleted === "true") {
              if (Subject !== subject) {
                event.completed({
                  allowEvent: false,
                  errorMessage:
                    "Title has been changed, Click Save in the taskpane before sending",
                });
              } else if (StartTime !== startTime) {
                event.completed({
                  allowEvent: false,
                  errorMessage:
                    "Start Date or Time has been changed, Click Save in the taskpane before sending",
                });
              } else if (EndTime !== endTime) {
                event.completed({
                  allowEvent: false,
                  errorMessage:
                    "End Date or Time has been changed, Click Save in the taskpane before sending",
                });
              } else if (Timezone !== timezone) {
                event.completed({
                  allowEvent: false,
                  errorMessage:
                    "TimeZone has been changed, Click Save in the taskpane before sending",
                });
              } else if (Location !== location) {
                event.completed({
                  allowEvent: false,
                  errorMessage:
                    "Location has been changed, Click Save in the taskpane before sending",
                });
              } else if (
                AttachmentCount !== attachmentCount ||
                AttachedFiles !== attachedFiles
              ) {
                event.completed({
                  allowEvent: false,
                  errorMessage:
                    "Attachment has been changed, Click Save in the taskpane before sending",
                });
              } else {
                event.completed({ allowEvent: true });
              }
            } else {
              event.completed({
                allowEvent: false,
                errorMessage:
                  "There are unsaved changes. Click 'Save' in the taskpane before sending.",
              });
            }
          });
        } else {
          console.log("There are no categories assigned to this item.");
          event.completed({ allowEvent: true });
        }
      } else {
        console.error(asyncResult.error);
        event.completed({ allowEvent: true });
      }
    } else {
      event.completed({ allowEvent: true });
    }
  });
}

function getFormattedDate(dateValue) {
  let timezoneRegex = /\(([^)]+)\)/;
  let timezoneID = timezoneRegex.exec(dateValue)[1];
  let dateISO = new Date(dateValue).toISOString();
  var formattedDateISO = new Date(dateISO);
  var formattedDateTime = new Date(
    formattedDateISO.getTime() - formattedDateISO.getTimezoneOffset() * 60000
  );
  let formattedDateTimeIS0 = new Date(formattedDateTime).toISOString();
  return {
    timezoneID: timezoneID,
    formattedDate: formattedDateTimeIS0.substring(0, 19),
  };
}

function onAppointmentChangeHandler(event) {
  //To be decided
  //console.log("Change detected", event);
  // Office.context.mailbox.item.categories.getAsync(async function (asyncResult) {
  //   if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
  //     const categories = asyncResult.value;
  //     console.log("Categories", categories);
  //     console.log("jbsdbjsd",categories.some((r) => r.displayName.includes("Synched")));
  //     if (categories && categories.length > 0) {
  //       if (
  //         categories.some((r) => r.displayName.includes("Nasdaq CMS Meetings"))
  //       ) {
  //         Office.context.mailbox.item.loadCustomPropertiesAsync(async function (
  //           asyncResult
  //         ) {
  //           const customProps = asyncResult.value;
  //           const eventId = customProps.get("eventID");
  //               console.log("Categories", categories);
  //               if (categories && categories.length > 0) {
  //                 // Grab the first category assigned to this item.
  //                 const categoryToRemove = [categories[1].displayName];
  //                 console.log("Category to be removed",categoryToRemove);
  //                 Office.context.mailbox.item.categories.removeAsync(categoryToRemove, function(asyncResult) {
  //                   if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
  //                     console.log(`Successfully unassigned category '${categoryToRemove}' from this item.`);
  //                   } else {
  //                     console.log("categories.removeAsync call failed with error: " + asyncResult.error.message);
  //                   }
  //                 });
  //               } else {
  //                 console.log("There are no categories assigned to this item.");
  //               }
  //           // customProps.set("synched", "false");
  //           // customProps.saveAsync(async function (asyncResult) {
  //           //   if (asyncResult.status === Office.AsyncResultStatus.Failed) {
  //           //     console.error(asyncResult.error.message);
  //           //   } else {
  //           //     console.log("Synched set as false");
  //           //   }
  //           // });
  //         });
  //       } else {
  //         console.log("There are no categories assigned to this item.");
  //         //event.completed({ allowEvent: true });
  //       }
  //     } else {
  //       console.error(asyncResult.error);
  //       //event.completed({ allowEvent: true });
  //     }
  //   } else {
  //     event.completed({ allowEvent: true });
  //   }
  // });
}

function onMessageChangeHandler(event) {
  //To be decided
}

function onMessageSendHandler(event) {
  Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
    const props = result.value;
    const categories = props.get("categories");
    const synched = props.get("synched");
    const subject = props.get("subject");
    const attachmentCount = props.get("attachmentCount");
    const attachedFiles = props.get("attachedFiles");
    if (categories === "Nasdaq CMS Meetings") {
      let mailboxItem = Office.context.mailbox.item;

      let AttachmentCount = "";
      let AttachedFiles = "";
      let Subject = "";

      mailboxItem.getAttachmentsAsync(function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error(result.error.message);
        } else {
          AttachmentCount = result.value.length.toString();
          if (result.value.length > 0) {
            let fileNameString = "";
            result.value.forEach((file) => {
              fileNameString = fileNameString + file.name;
            });
            AttachedFiles = fileNameString;
          }
        }
      });

      mailboxItem.subject.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to get subject", asyncResult.error.message);
        } else {
          
          Subject = asyncResult.value;

          if (Subject === "") {
            event.completed({
              allowEvent: false,
              errorMessage: "Title cannot be blank. Please add ",
            });
          } else if (synched === "true") {
            if (Subject !== subject) {
              event.completed({
                allowEvent: false,
                errorMessage:
                  "Title has been changed, Click Save in the taskpane before sending",
              });
            } else if (
              AttachmentCount !== attachmentCount ||
              AttachedFiles !== attachedFiles
            ) {
              event.completed({
                allowEvent: false,
                errorMessage:
                  "Attachment has been changed, Click Save in the taskpane before sending",
              });
            } else {
              event.completed({ allowEvent: true });
            }
          } else {
            event.completed({
              allowEvent: false,
              errorMessage:
                "There are unsaved changes. Click 'Save' in the taskpane before sending.",
            });
          }

        }
      });
    } else {
      console.log("Not a cms email");
      event.completed({ allowEvent: true });
    }
  });
}

Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
Office.actions.associate(
  "onAppointmentTimeChangedHandler",
  onAppointmentChangeHandler
);
Office.actions.associate(
  "onAppointmentRecurrenceChangedHandler",
  onAppointmentChangeHandler
);
Office.actions.associate(
  "onAppointmentAttachmentsChangedHandler",
  onAppointmentChangeHandler
);
Office.actions.associate(
  "onAppointmentAttendeesChangedHandler",
  onAppointmentChangeHandler
);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate(
  "onMessageAttachmentsChangedHandler",
  onMessageChangeHandler
);
