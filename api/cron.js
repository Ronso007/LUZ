// const { google } = require("googleapis");
// var fs = require("fs");

// export default async function downloadLuz(req, res) {
//   const drive = google.drive({ version: "v3", auth: global.oauth2Client });
//   const fileId = "1x6PjiaUTt8u5E3NoN6b8ogYcVpwyEnAS";
//   const dest = fs.createWriteStream("LUZ2.xlsx");
//   try {
//     const file = await drive.files.get(
//       {
//         fileId: fileId,
//         alt: "media",
//       },
//       { responseType: "arraybuffer" }
//     );
//     // file.data.on("end", () => console.log("onCompleted"));
//     //file.data.pipe(dest);
//     fs.writeFileSync("LUZ2.xlsx", Buffer.from(file.data));
//     //return file.status;
//   } catch (err) {
//     // TODO(developer) - Handle error
//     throw err;
//   }
//   console.log("Success excel file");
// }
