// script/clientdata.js
//
// One-shot seeder: inserts all clients from the spreadsheet into MongoDB.
// Run from the backend folder:
//
//     node script/clientdata.js
//
// Uses the same MONGO_URI as the server (read from backend/.env).
// Deletes existing clients first so re-running is idempotent.
//
// To skip the wipe, comment out the `deleteMany` line below.

import dotenv from "dotenv";
import mongoose from "mongoose";

dotenv.config();

// ------------------------------------------------------------------
// Schema — mirror of the one in server.js so the fields line up.
// Declared here rather than imported so this script is fully
// self-contained (you can run it without touching server.js).
// ------------------------------------------------------------------
const clientSchema = new mongoose.Schema({
  partyName: { type: String, required: true, trim: true },
  firmName: { type: String, trim: true, default: "" },
  userId: { type: String, trim: true, default: "" },
  password: { type: String, default: "" },
  mobileNo: { type: String, trim: true, default: "" },
  emailId: { type: String, trim: true, lowercase: true, default: "" },
  licenseNumber: { type: String, trim: true, default: "" },
  licenseType: { type: String, trim: true, default: "" },
  clientNumber: { type: String, trim: true, default: "" },
  designation: { type: String, trim: true, default: "" },
  kob: { type: String, trim: true, default: "" },

  expiredDate: { type: Date, required: true },
  expiredDateFormat: { type: String, default: "dd-mm-yyyy" },

  lastRenewedAt: { type: Date, default: null },
  renewed: { type: Boolean, default: false },
  dismissedOn: { type: String, default: null },
  remindersSent: { type: [Number], default: [] },

  createdAt: { type: Date, default: Date.now },
  updatedAt: { type: Date, default: Date.now }
});

const Client = mongoose.model("Client", clientSchema);

// ------------------------------------------------------------------
// Helpers
// ------------------------------------------------------------------
// "20-09-2026" → Date at LOCAL midnight of that calendar day.
// Matches parseDate() in server.js so daysUntil() returns the
// same number as the dashboard / extension.
function dmy(str) {
  const [d, m, y] = str.split("-").map(Number);
  return new Date(y, m - 1, d, 0, 0, 0, 0);
}

function s(v) {
  if (v === undefined || v === null) return "";
  const t = String(v).trim();
  return t === "-" ? "" : t;
}

// ------------------------------------------------------------------
// Data
// ------------------------------------------------------------------
const clients = [
  { partyName: "SUNITA GURUNG",                firmName: "SAUGAAT ENTERPRISE",                        userId: "12321002000162", password: "Kuberji@1234",  mobileNo: "9832949193",             emailId: "itrcheck02@gmail.com",              licenseNumber: "12321002000162", licenseType: "State License",   clientNumber: "9474586243",   designation: "PROPRIETOR", expiredDate: dmy("20-09-2026"), kob: "Retailer" },
  { partyName: "SURENDRA PRASAD/AKASH PRASAD", firmName: "MAA KAMAKHYA ENTERPRISE PARTNER",           userId: "12823009000285", password: "Kuberji@123",   mobileNo: "9832949193",             emailId: "itrcheck023@gmail.com",             licenseNumber: "12823009000285", licenseType: "State License",   clientNumber: "9402500000",   designation: "PARTNER",    expiredDate: dmy("06-11-2026"), kob: "Wholesaler" },
  { partyName: "BIDYA SAGAR PRASAD",           firmName: "BIDYA SAGAR PRASAD",                        userId: "12822024000062", password: "Kuberji@1234",  mobileNo: "9832093139",             emailId: "bidyasagarp5@gmail.com",            licenseNumber: "12822024000062", licenseType: "State License",   clientNumber: "9832012359",   designation: "PROPRIETOR", expiredDate: dmy("23-11-2026"), kob: "Wholesaler" },
  { partyName: "SUNIL KUMAR",                  firmName: "SHREYANSH ENTERPRISE",                      userId: "12825006000626", password: "",              mobileNo: "9870323733",             emailId: "itrcheck01@gmail.com",              licenseNumber: "12825006000626", licenseType: "State License",   clientNumber: "9870323733",   designation: "PROPRIETOR", expiredDate: dmy("30-11-2026"), kob: "Retailer" },
  { partyName: "JAGADISH SIGCHI",              firmName: "LAXMI STORE",                               userId: "12822006000475", password: "Kuberji@123",   mobileNo: "9832093139",             emailId: "itrcheck042@gmail.com",             licenseNumber: "12822006000475", licenseType: "State License",   clientNumber: "9593677078",   designation: "PROPRIETOR", expiredDate: dmy("02-12-2026"), kob: "Distributor/Wholesaler" },
  { partyName: "MITHUN KANTI MUHURI",          firmName: "MAA DURGA AGENCY",                          userId: "12824006000490", password: "Abcd1234@",     mobileNo: "9832381113",             emailId: "itrfiling065@gmail.com",            licenseNumber: "12824006000490", licenseType: "State License",   clientNumber: "9832381113",   designation: "PROPRIETOR", expiredDate: dmy("09-12-2026"), kob: "Wholesaler" },
  { partyName: "ANAND JAISWAL",                firmName: "THE SUPERFOOD COMPANY",                     userId: "12824999001184", password: "Kuberji@123",   mobileNo: "8011196093",             emailId: "itrcheck050@gmail.com",             licenseNumber: "12824999001184", licenseType: "Central License", clientNumber: "8011196093",   designation: "PROPRIETOR", expiredDate: dmy("09-12-2026"), kob: "Retailer/Wholesaler" },
  { partyName: "JITENDRA NATH ROY",            firmName: "KISHOLAY TEA LLP",                          userId: "12824006000501", password: "Kuberji@123",   mobileNo: "7001973251",             emailId: "itrcheck072@gmail.com",             licenseNumber: "12824006000501", licenseType: "State License",   clientNumber: "7001973251",   designation: "LLP/PARTNER", expiredDate: dmy("12-12-2026"), kob: "Wholesaler" },
  { partyName: "ANUP GUPTA",                   firmName: "AR ENTERPRISE",                             userId: "12819006000216", password: "Kuberji@021",   mobileNo: "9832093139",             emailId: "itrcheck036@gmail.com",             licenseNumber: "12819006000216", licenseType: "State License",   clientNumber: "9932476294",   designation: "PROPRIETOR", expiredDate: dmy("31-12-2026"), kob: "Wholesaler" },
  { partyName: "MANISH RAI",                   firmName: "MANISH RAI",                                userId: "12819006000231", password: "Kuberji@021",   mobileNo: "9832093139",             emailId: "itrcheck035@gmail.com",             licenseNumber: "12819006000231", licenseType: "State License",   clientNumber: "8768934299, 9734914254", designation: "PROPRIETOR", expiredDate: dmy("31-12-2026"), kob: "Retailer" },
  { partyName: "RATAN KUMAR CHOUDHARY",        firmName: "NEW  MAHARAJA",                             userId: "12819006000399", password: "Kuberji@1234",  mobileNo: "9832093139",             emailId: "itrcheck011@gmail.com",             licenseNumber: "12819006000399", licenseType: "State License",   clientNumber: "9931636783",   designation: "PROPRIETOR", expiredDate: dmy("31-12-2026"), kob: "Retailer" },
  { partyName: "VIJAY KUMAR RAY/BISWADEEP PAUL", firmName: "SHUANVI ESSENTIAL",                       userId: "12825006000034", password: "Abcd1234@",     mobileNo: "9832399900",             emailId: "vijayrayy@gmail.com",               licenseNumber: "12825006000034", licenseType: "State License",   clientNumber: "9832399900",   designation: "PARTNER",    expiredDate: dmy("15-01-2027"), kob: "Wholesaler" },
  { partyName: "BINAY KUMAR GUPTA",            firmName: "BAJRANG ENTERPRISES",                       userId: "12826006000028", password: "Abcd1234@",     mobileNo: "7602567999",             emailId: "binaygupta283@gmail.com",           licenseNumber: "12826006000028", licenseType: "State License",   clientNumber: "7602567999",   designation: "PROPRIETOR", expiredDate: dmy("16-01-2027"), kob: "Distributor" },
  { partyName: "PUSHPARANI THAKUR",            firmName: "SHRUTI ENTERPRISE",                         userId: "12826006000029", password: "Abcd1234@",     mobileNo: "9832028328",             emailId: "thakurshawan12@gmail.com",          licenseNumber: "12826006000029", licenseType: "State License",   clientNumber: "9832028328",   designation: "PROPRIETOR", expiredDate: dmy("16-01-2027"), kob: "Wholesaler/Distributor" },
  { partyName: "MANOJ KUMAR GUPTA",            firmName: "MANOJ KUMAR GUPTA",                         userId: "12823006000026", password: "Kuberji@1234",  mobileNo: "9635061174",             emailId: "itrcheck020@gmail.com",             licenseNumber: "12823006000026", licenseType: "State License",   clientNumber: "9635061174",   designation: "PROPRIETOR", expiredDate: dmy("17-01-2027"), kob: "Retailer/Wholesaler" },
  { partyName: "ASHOK KUMAR AGARWAL",          firmName: "DURGA DEPARTMENTAL STORE",                  userId: "12821006000491", password: "Kuberji@021",   mobileNo: "8250843932",             emailId: "itrcheck019@gmail.com",             licenseNumber: "12821006000491", licenseType: "State License",   clientNumber: "8250843932",   designation: "PROPRIETOR", expiredDate: dmy("19-01-2027"), kob: "Retailer" },
  { partyName: "UDAY PRAKASH CHOUDHARY",       firmName: "DARJEELING FRESH TEA",                      userId: "12825006000048", password: "Abcd1234@",     mobileNo: "7866960010",             emailId: "udaychoudhary01@gmail.com",         licenseNumber: "12825006000048", licenseType: "State License",   clientNumber: "7866960010",   designation: "PROPRIETOR", expiredDate: dmy("22-01-2027"), kob: "Retailer" },
  { partyName: "NARENDRA KUMAR JAIN",          firmName: "JAIN CORPORATION",                          userId: "12821006000038", password: "Abcd1234@",     mobileNo: "9785920018",             emailId: "vikasioplus@gmail.com",             licenseNumber: "12821006000038", licenseType: "State License",   clientNumber: "9785920018",   designation: "PROPRIETOR", expiredDate: dmy("28-01-2027"), kob: "Retail/Wholesaler" },
  { partyName: "ABHISHEK GUPTA",               firmName: "KHATU SHYAM TRADING",                       userId: "12825006000069", password: "Kuberji@123",   mobileNo: "8181818340",             emailId: "itrcheck025@gmail.com",             licenseNumber: "12825006000069", licenseType: "State License",   clientNumber: "8181818340",   designation: "PROPRIETOR", expiredDate: dmy("30-01-2027"), kob: "Wholesaler" },
  { partyName: "UDAY PRAKASH CHOUDHARY",       firmName: "UD RESTAURANT",                             userId: "12824009000020", password: "Kuberji@123",   mobileNo: "9832949193",             emailId: "itrcheck030@gmail.com",             licenseNumber: "12824009000020", licenseType: "State License",   clientNumber: "7866960010",   designation: "PROPRIETOR", expiredDate: dmy("31-01-2027"), kob: "Retailer" },
  { partyName: "SHIWJEE PRASAD",               firmName: "PARMESHWAR PRASAD SHIWJEE PRASAD",          userId: "10021031000108", password: "Abcd1234@",     mobileNo: "8927905581",             emailId: "itrcheck012@gmail.com",             licenseNumber: "10021031000108", licenseType: "Central License", clientNumber: "8250607714",   designation: "PARTNER",    expiredDate: dmy("02-02-2027"), kob: "Food Service- Restaurant" },
  { partyName: "SUMAN SARKAR",                 firmName: "SARKAR AGENCY",                             userId: "12826009000049", password: "Abcd1234@",     mobileNo: "7001445522",             emailId: "suman7.2014@gmail.com",             licenseNumber: "12826009000049", licenseType: "State License",   clientNumber: "7001445522",   designation: "PROPRIETOR", expiredDate: dmy("27-02-2027"), kob: "Importer/Wholesaler/Distributor/Retailer/Supplier/Marketer/Exporter" },
  { partyName: "SUJIT KUMAR SAHA",             firmName: "SABITA TRADERS",                            userId: "12822006000114", password: "Kuberji@1234",  mobileNo: "9832044471",             emailId: "itrcheck030@gmail.com",             licenseNumber: "12822006000114", licenseType: "State License",   clientNumber: "9832044471",   designation: "PROPRIETOR", expiredDate: dmy("13-03-2027"), kob: "Retailer" },
  { partyName: "AJIT KUMAR PRASAD",            firmName: "AJIT KUMAR PRASAD",                         userId: "12324001000025", password: "Kuberji@123",   mobileNo: "9832949193",             emailId: "itrcheck028@gmail.com",             licenseNumber: "12324001000025", licenseType: "State License",   clientNumber: "9434485180",   designation: "PROPRIETOR", expiredDate: dmy("25-03-2027"), kob: "Wholesaler" },
  { partyName: "JAIDEEP BANERJEE",             firmName: "ARIXA HEALTHCARE PRIVATE LIMITED",          userId: "12826006000151", password: "Abcd1234@",     mobileNo: "9476391310",             emailId: "arixahealth@gmail.com",             licenseNumber: "12826006000151", licenseType: "State License",   clientNumber: "9476391310",   designation: "COMPANY",    expiredDate: dmy("25-03-2027"), kob: "Retailer" },
  { partyName: "JAIDEEP BANERJEE",             firmName: "NEXA LIFE PHARMACEUTICALS PRIVATE LIMITED", userId: "12826009000105", password: "Abcd1234@",     mobileNo: "9476391310",             emailId: "joydeep.san@gmail.com",             licenseNumber: "12826009000105", licenseType: "State License",   clientNumber: "9476391310",   designation: "COMPANY",    expiredDate: dmy("06-05-2027"), kob: "Retailer/Wholesaler" },
  { partyName: "RAJIV PRASAD GUPTA",           firmName: "PANKAJ KUMAR AKASH KUMAR",                  userId: "12822006000221", password: "Abcd1234@",     mobileNo: "9832364122",             emailId: "ankitslg.agarwal15@gmail.com",      licenseNumber: "12822006000221", licenseType: "State License",   clientNumber: "9832364122",   designation: "PROPRIETOR", expiredDate: dmy("24-05-2027"), kob: "Wholesaler" },
  { partyName: "RAJEEV KUMAR SURANA",          firmName: "BALAJI IMPEX",                              userId: "22826084000499", password: "Abcd1234@",     mobileNo: "9212360986",             emailId: "cacslalitgupta1@gmail.com",         licenseNumber: "22826084000499", licenseType: "Registration",    clientNumber: "9212360986",   designation: "PROPRIETOR", expiredDate: dmy("29-06-2027"), kob: "Retailer" },
  { partyName: "PRIYANKA SEKHSARIA",           firmName: "BHARAT MULTIEXIM",                          userId: "12821999000052", password: "Kuberji@021",   mobileNo: "9331069728",             emailId: "itrcheck013@gmail.com",             licenseNumber: "12821999000052", licenseType: "Central License", clientNumber: "8250607714",   designation: "PARTNER",    expiredDate: dmy("14-07-2027"), kob: "Wholesaler" },
  { partyName: "RAJEEV KUMAR SURANA",          firmName: "AMBEY TRADING",                             userId: "13326005000469", password: "Abcd1234@",     mobileNo: "9212360986",             emailId: "cacslalitgupta1@gmail.com",         licenseNumber: "13326005000469", licenseType: "State License",   clientNumber: "9212360986",   designation: "PROPRIETOR", expiredDate: dmy("20-07-2027"), kob: "Wholesaler" },
  { partyName: "MIRADEVI GUPTA",               firmName: "M/S M D ENTERPRISES",                       userId: "12822006000326", password: "Abcd1234@",     mobileNo: "9635061174",             emailId: "miradevigupta2022@gmail.com",       licenseNumber: "12822006000326", licenseType: "State License",   clientNumber: "9635061174",   designation: "PROPRIETOR", expiredDate: dmy("26-07-2027"), kob: "Supplier" },
  { partyName: "KARMA CHIMI BHUTIA",           firmName: "KARMA CHIMI BHUTIA",                        userId: "12325001000071", password: "Abcd1234@@@",   mobileNo: "9832703802",             emailId: "karmachimibhutia8@gmail.com",       licenseNumber: "12325001000071", licenseType: "State License",   clientNumber: "9832703802",   designation: "PROPRIETOR", expiredDate: dmy("30-07-2027"), kob: "Wholesaler" },
  { partyName: "RINA DEVI PRASAD",             firmName: "SAFARI RESTAURANT AND HOMESTAY",            userId: "22826191000310", password: "Abcd1234@",     mobileNo: "9800500890",             emailId: "",                                  licenseNumber: "22826191000310", licenseType: "Registration",    clientNumber: "9800500890",   designation: "PROPRIETOR", expiredDate: dmy("02-08-2027"), kob: "Hotel" },
  { partyName: "RENUKA DEVI DAHAL",            firmName: "M/S YAVNIKA STORES",                        userId: "12320002000047", password: "Abcd1234@",     mobileNo: "7908099053",             emailId: "thapasanjay8700@gmail.com",         licenseNumber: "12320002000047", licenseType: "State License",   clientNumber: "7908099053",   designation: "PROPRIETOR", expiredDate: dmy("10-08-2027"), kob: "Distributor" },
  { partyName: "DHANANJAY PRASAD JAISWAL",     firmName: "JAISWAL AGRO CROPS PVT LTD",                userId: "12822018000123", password: "Kuberji@1234",  mobileNo: "9832949193",             emailId: "itrcheck06@gmail.com",              licenseNumber: "12822018000123", licenseType: "State License",   clientNumber: "8011196093",   designation: "COMPANY",    expiredDate: dmy("12-08-2027"), kob: "Retailer" },
  { partyName: "VIKRAM KUMAR GUPTA",           firmName: "TIRUPATI TEA CARRIER",                      userId: "12821009000232", password: "Kuberji@1234",  mobileNo: "9002221225",             emailId: "itrcheck010@gmail.com",             licenseNumber: "12821009000232", licenseType: "State License",   clientNumber: "",             designation: "PROPRIETOR", expiredDate: dmy("06-09-2027"), kob: "Wholesaler" },
  { partyName: "PRADEEP KUMAR GUPTA",          firmName: "AMEYAA ENTERPRISES",                        userId: "12824006000328", password: "Kuberji@123",   mobileNo: "9832949193",             emailId: "itrcheck020@gmail.com",             licenseNumber: "12824006000328", licenseType: "State License",   clientNumber: "7800851641",   designation: "PROPRIETOR", expiredDate: dmy("08-09-2027"), kob: "Supplier" },
  { partyName: "SANJAY KUMAR SARAOGI",         firmName: "ARIHANT TRADERS",                           userId: "12819006000456", password: "Kuberji@1234",  mobileNo: "9832066751",             emailId: "itrcheck036@gmail.com",             licenseNumber: "12819006000456", licenseType: "State License",   clientNumber: "825070730",    designation: "PROPRIETOR", expiredDate: dmy("22-10-2027"), kob: "Wholesaler" },
  { partyName: "ROSHAN LAL GUPTA",             firmName: "ROSHAN TRADERS",                            userId: "12825006000041", password: "Abcd1234@",     mobileNo: "7001029692",             emailId: "itrcheck01@gmail.com",              licenseNumber: "12825006000041", licenseType: "State License",   clientNumber: "7001029692",   designation: "PROPRIETOR", expiredDate: dmy("20-01-2028"), kob: "Wholesaler" },
  { partyName: "SAMIR PRASAD",                 firmName: "MAHADEV TRADING",                           userId: "12823006000319", password: "Abcd1234@",     mobileNo: "8250707301",             emailId: "samirprasad032@gmail.com",          licenseNumber: "12823006000319", licenseType: "State License",   clientNumber: "8250707301",   designation: "PROPRIETOR", expiredDate: dmy("12-07-2028"), kob: "Wholesaler" },
  { partyName: "SANJAY KUMAR GUPTA",           firmName: "SANJAY KUMAR RAJEEV KUMAR",                 userId: "12819006000400", password: "Kuberji@321",   mobileNo: "9832093139",             emailId: "itrcheck019@gmail.com",             licenseNumber: "12819006000400", licenseType: "State License",   clientNumber: "9832041211",   designation: "PROPRIETOR", expiredDate: dmy("31-12-2028"), kob: "Distributor" },
  { partyName: "VIVEK OSWAL",                  firmName: "JAIN HOMESTAY",                             userId: "22823084000870", password: "Abcd1234@",     mobileNo: "9382256884",             emailId: "jainhomestay@gmail.com",            licenseNumber: "22823084000870", licenseType: "Registration",    clientNumber: "9382256884",   designation: "PROPRIETOR", expiredDate: dmy("16-07-2029"), kob: "Retailer" },
  { partyName: "RAJ KUMAR JAIN",               firmName: "RAJ KUMAR JAIN",                            userId: "12822006000279", password: "Kuberji@123",   mobileNo: "9832949193",             emailId: "itrcheck025@gmail.com",             licenseNumber: "12822006000279", licenseType: "State License",   clientNumber: "9775897971",   designation: "PROPRIETOR", expiredDate: dmy("21-06-2031"), kob: "Retailer" },
  { partyName: "PANKAJ KUMAR JAIN",            firmName: "MAHI ENTERPRISES",                          userId: "12822006000325", password: "Kuberji@12345", mobileNo: "9832949193",             emailId: "itrcheck020@gmail.com",             licenseNumber: "12822006000325", licenseType: "State License",   clientNumber: "9775577781",   designation: "PROPRIETOR", expiredDate: dmy("11-07-2031"), kob: "Wholesaler" },
  { partyName: "SHYAM SUNDAR AGARWAL",         firmName: "SHYAM SUNDAR BIMAL KUMAR",                  userId: "12821006000272", password: "Abcd1234@",     mobileNo: "9434045153",             emailId: "sunilag_ssbk@yahoo.in",             licenseNumber: "12821006000272", licenseType: "State License",   clientNumber: "9434045153",   designation: "PROPRIETOR", expiredDate: dmy("01-08-2031"), kob: "Wholesaler" }
];

// ------------------------------------------------------------------
// Transform → MongoDB documents
// ------------------------------------------------------------------
const docs = clients.map(c => ({
  partyName:         s(c.partyName),
  firmName:          s(c.firmName),
  userId:            s(c.userId),
  password:          s(c.password),
  mobileNo:          s(c.mobileNo),
  emailId:           s(c.emailId).toLowerCase(),
  licenseNumber:     s(c.licenseNumber),
  licenseType:       s(c.licenseType),
  clientNumber:      s(c.clientNumber),
  designation:       s(c.designation),
  kob:               s(c.kob),
  expiredDate:       c.expiredDate,
  expiredDateFormat: "dd-mm-yyyy",
  lastRenewedAt:     null,
  renewed:           false,
  dismissedOn:       null,
  remindersSent:     [],
  createdAt:         new Date(),
  updatedAt:         new Date()
}));

// ------------------------------------------------------------------
// Main
// ------------------------------------------------------------------
async function main() {
  const uri = process.env.MONGO_URI;
  if (!uri) {
    console.error("MONGO_URI is not set. Put it in backend/.env or export it in the shell.");
    process.exit(1);
  }

  console.log("Connecting to MongoDB…");
  await mongoose.connect(uri);
  console.log("Connected.");

  // ── OPTIONAL: wipe existing clients so re-running is idempotent.
  // Comment out this block if you want to *append* instead.
  const removed = await Client.deleteMany({});
  console.log(`Removed ${removed.deletedCount} existing client(s).`);
  // ────────────────────────────────────────────────────────────────

  const inserted = await Client.insertMany(docs);
  console.log(`✅ Inserted ${inserted.length} client(s).`);
  console.log(`📊 Total clients in collection now: ${await Client.countDocuments()}`);

  await mongoose.disconnect();
  console.log("Disconnected. Done.");
}

main().catch(err => {
  console.error("Seed failed:", err);
  process.exit(1);
});