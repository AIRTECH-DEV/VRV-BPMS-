// function handleFormSubmit(e) {
//   const sheetName = e.range.getSheet().getName();

//   // 1️⃣ Always generate Complaint ID if applicable
//   assignOrderId(e);

//   // 2️⃣ Route based on the sheet name
//   // const HANDLERS = {
//   //   "Complaint Entry": handleComplaintForm, // complaint form
//   //   "Complaint Report": handleComplaintReportForm,    // report upload form
//   //   // "Form Responses 3": handlePaymentForm,   // payment follow-up form
//   // };

//   const handler = HANDLERS[sheetName];
//   if (handler) {
//     handler(e);
//   } else {
//     Logger.log(`⚠️ No handler defined for sheet: ${sheetName}`);
//   }
// }