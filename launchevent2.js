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

          const item = Office.context.mailbox.item;
          item.loadCustomPropertiesAsync((result) => {
            const props = result.value;
            const synched = props.get("synched");
            const subject = props.get("subject");
            const deleted = props.get("deleted");
            // console.log("Synched", synched);
            // console.log("Subject", subject);
            // console.log("Deleted", deleted);
            let mailboxItem = Office.context.mailbox.item;
            mailboxItem.subject.getAsync((asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(
                  "Failed to get subject",
                  asyncResult.error.message
                );
              } else {
                // Successfully got the subject, display it.
                //console.log("SUBJECT", asyncResult.value);
                if (asyncResult.value === "") {
                  console.log("111");
                  event.completed({
                    allowEvent: false,
                    errorMessage: "Subject can't be blank",
                  });
                } else if (
                  subject !== undefined &&
                  asyncResult.value !== subject
                ) {
                  console.log("222");
                  event.completed({
                    allowEvent: false,
                    errorMessage:
                      "Subject has been changed,Click Save before sending in the taskpane.",
                  });
                } else if (synched === "true" || deleted === "true") {
                  if (asyncResult.value === subject) {
                    // const propertyName = "synched".val();
                    props.remove(synched);
                    props.remove(subject);
                    props.remove(deleted);
                    //props.remove(eventUpdated);
                    console.log(`Custom property "${synched}" removed.`);
                    props.saveAsync(async function (asyncResult) {
                      if (
                        asyncResult.status === Office.AsyncResultStatus.Failed
                      ) {
                        console.error(asyncResult.error.message);
                      } else {
                        //console.log("Properties Saved");

                         event.completed({ allowEvent: true });

                        //if (categories && categories.length > 0) {
                          // Grab the first category assigned to this item.
                          // const categoryToRemove = [categories[1].displayName];
                          // console.log(
                          //   "Category to be removed",
                          //   categoryToRemove
                          // );
                          // Office.context.mailbox.item.categories.removeAsync(
                          //   categoryToRemove,
                          //   function (asyncResult) {
                          //     if (
                          //       asyncResult.status ===
                          //       Office.AsyncResultStatus.Succeeded
                          //     ) {
                          //       console.log(
                          //         `Successfully unassigned category '${categoryToRemove}' from this item.`
                          //       );
                          
                          //     } else {
                          //       console.log(
                          //         "categories.removeAsync call failed with error: " +
                          //           asyncResult.error.message
                          //       );
                          //     }
                          //   }
                          // );
                        // } else {
                        //   console.log(
                        //     "There are no categories assigned to this item."
                        //   );
                        // }

                        // Async call to save custom properties completed.
                        // Proceed to do the appropriate for your add-in.
                      }
                    });
                  } else {
                    // let platform = Office.context.diagnostics.platform;
                    // console.log("Platform is", platform);
                    // if (platform === Office.PlatformType.OfficeOnline) {
                      event.completed({
                        allowEvent: false,
                        errorMessage:
                          "There are unsaved changes. Click 'Save' in the taskpane before sending.",
                      });
                    // } else {
                    //   event.completed({ allowEvent: true });
                    // }
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
    if (categories === "Nasdaq CMS Meetings") {
      let mailboxItem = Office.context.mailbox.item;
      mailboxItem.subject.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to get subject", asyncResult.error.message);
        } else {
          // Successfully got the subject, display it.
          if (asyncResult.value === "") {
            //console.log("111");
            event.completed({
              allowEvent: false,
              errorMessage: "Subject can't be blank",
            });
          } else if (subject !== undefined && asyncResult.value !== subject) {
            event.completed({
              allowEvent: false,
              errorMessage:
                "Subject has been changed,Click Save before sending in the taskpane.",
            });
          } else if (synched === "true") {
            if (asyncResult.value === subject) {
              props.remove(synched);
              props.remove(subject);
              props.saveAsync(async function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  console.error(asyncResult.error.message);
                } else {
                  event.completed({ allowEvent: true });
                }
              });
            } else {
              let platform = Office.context.diagnostics.platform;
              console.log("Platform is", platform);
              if (platform === Office.PlatformType.OfficeOnline) {
                event.completed({
                  allowEvent: false,
                  errorMessage:
                    "Subject has been changed,Click Save before sending in the taskpane.",
                });
              } else {
                event.completed({ allowEvent: true });
              }
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
