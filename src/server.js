import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import mongoose from "mongoose";
import XLSX from "xlsx";
import nodemailer from "nodemailer";
import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";

dotenv.config();

const app = express();

const port = process.env.PORT || 5000;

app.use(
  cors({
    origin: process.env.CORS_ORIGIN?.split(",") || "*",
    exposedHeaders: ["Content-Disposition"],   // 👈 ADD THIS LINE
  }),
);

app.use(express.json());

// ==================================================================
// Auth config
// ==================================================================

const JWT_SECRET = process.env.JWT_SECRET || "please-change-this-secret-in-env";

const TOKEN_EXPIRES_IN = process.env.TOKEN_EXPIRES_IN || "12h";

const DEFAULT_ADMIN_USERNAME = process.env.ADMIN_USERNAME || "admin";

const DEFAULT_ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "Admin@123";

if (!process.env.JWT_SECRET) {
  console.warn(
    "[auth] WARNING: JWT_SECRET is not set in .env — using an insecure default. " +
      "Set JWT_SECRET in backend/.env before using this in production.",
  );
}

// ==================================================================
// Schemas
// ==================================================================

const userSchema = new mongoose.Schema({
  username: {
    type: String,
    required: true,
    unique: true,
    trim: true,
    lowercase: true,
  },

  passwordHash: {
    type: String,
    required: true,
  },

  createdAt: {
    type: Date,
    default: Date.now,
  },
});

const User = mongoose.model("User", userSchema);

// ------------------------------------------------------------------
// Client record
// ------------------------------------------------------------------

const clientSchema = new mongoose.Schema({
  partyName: {
    type: String,
    required: true,
    trim: true,
  },

  firmName: {
    type: String,
    trim: true,
    default: "",
  },

  userId: {
    type: String,
    trim: true,
    default: "",
  },

  password: {
    type: String,
    default: "",
  },

  mobileNo: {
    type: String,
    trim: true,
    default: "",
  },

  emailId: {
    type: String,
    trim: true,
    lowercase: true,
    default: "",
  },

  licenseNumber: {
    type: String,
    trim: true,
    default: "",
  },

  licenseType: {
    type: String,
    trim: true,
    default: "",
  },

  clientNumber: {
    type: String,
    trim: true,
    default: "",
  },

  designation: {
    type: String,
    trim: true,
    default: "",
  },

  kob: {
    type: String,
    trim: true,
    default: "",
  },

  expiredDate: {
    type: Date,
    required: true,
  },

  expiredDateFormat: {
    type: String,
    default: "dd-mm-yyyy",
  },

  lastRenewedAt: {
    type: Date,
    default: null,
  },

  renewed: {
    type: Boolean,
    default: false,
  },

  dismissedOn: {
    type: String,
    default: null,
  },

  remindersSent: {
    type: [Number],
    default: [],
  },

  createdAt: {
    type: Date,
    default: Date.now,
  },

  updatedAt: {
    type: Date,
    default: Date.now,
  },
});

const Client = mongoose.model("Client", clientSchema);

// ==================================================================
// Client options
// ==================================================================

const CLIENT_OPTIONS = {
  licenseType: [
    "State License",
    "Central License",
    "Registration",
    "Basic Registration",
  ],

  designation: [
    "PROPRIETOR",
    "PARTNER",
    "LLP/PARTNER",
    "DIRECTOR",
    "COMPANY",
    "AUTHORIZED SIGNATORY",
    "KARTA",
  ],

  kob: [
    "Retailer",
    "Wholesaler",
    "Distributor",
    "Distributor/Wholesaler",
    "Retailer/Wholesaler",
    "Retail/Wholesaler",
    "Manufacturer",
    "Importer",
    "Exporter",
    "Supplier",
    "Marketer",
    "Importer/Wholesaler/Distributor/Retailer/Supplier/Marketer/Exporter",
    "Food Service- Restaurant",
    "Hotel",
  ],
};

const DEFAULT_DATE_FORMAT = "dd-mm-yyyy";

// ==================================================================
// Date helpers
// ==================================================================

function atMidnight(d) {
  const x = d instanceof Date ? d : new Date(d);

  return new Date(x.getFullYear(), x.getMonth(), x.getDate(), 0, 0, 0, 0);
}

function daysUntil(date) {
  if (!date) return null;

  const today = atMidnight(new Date());

  const expiry = atMidnight(date);

  return Math.round((expiry - today) / 86400000);
}

function isValidDateParts(year, month, day) {
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    !Number.isInteger(day)
  ) {
    return false;
  }

  const d = new Date(year, month - 1, day);

  return (
    d.getFullYear() === year &&
    d.getMonth() === month - 1 &&
    d.getDate() === day
  );
}

function parseDate(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) {
      return null;
    }

    return new Date(
      value.getFullYear(),
      value.getMonth(),
      value.getDate(),
      0,
      0,
      0,
      0,
    );
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);

    if (!parsed || !isValidDateParts(parsed.y, parsed.m, parsed.d)) {
      return null;
    }

    return new Date(parsed.y, parsed.m - 1, parsed.d, 0, 0, 0, 0);
  }

  const text = String(value).trim();

  let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);

  if (match) {
    const [, y, m, d] = match.map(Number);

    return isValidDateParts(y, m, d) ? new Date(y, m - 1, d, 0, 0, 0, 0) : null;
  }

  match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);

  if (match) {
    let [, d, m, y] = match;

    if (y.length === 2) {
      y = `20${y}`;
    }

    d = Number(d);
    m = Number(m);
    y = Number(y);

    return isValidDateParts(y, m, d) ? new Date(y, m - 1, d, 0, 0, 0, 0) : null;
  }

  const date = new Date(text);

  return Number.isNaN(date.getTime())
    ? null
    : new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
}

function normalizeDateFormat(format) {
  if (!format || typeof format !== "string") {
    return DEFAULT_DATE_FORMAT;
  }

  const cleaned = format.trim().replace(/\\-/g, "-").replace(/\\\//g, "/");

  const lower = cleaned.toLowerCase();

  if (
    lower.includes("yy") &&
    (lower.includes("dd") || lower.includes("d")) &&
    (lower.includes("mm") || lower.includes("m"))
  ) {
    return cleaned;
  }

  return DEFAULT_DATE_FORMAT;
}

function formatDateByPattern(value, pattern = DEFAULT_DATE_FORMAT) {
  const date = parseDate(value);

  if (!date) return "";

  const d = String(date.getDate()).padStart(2, "0");

  const m = String(date.getMonth() + 1).padStart(2, "0");

  const y = String(date.getFullYear());

  const yy = y.slice(-2);

  let out = pattern;

  out = out.replace(/yyyy/gi, y).replace(/yy/gi, yy);

  out = out.replace(/dd/gi, d);

  out = out.replace(/mm/gi, m);

  out = out.replace(/d/g, String(date.getDate()));

  out = out.replace(/m/g, String(date.getMonth() + 1));

  return out;
}

function dateKey(date = new Date()) {
  const d = date instanceof Date ? date : new Date(date);

  const y = d.getFullYear();

  const m = String(d.getMonth() + 1).padStart(2, "0");

  const day = String(d.getDate()).padStart(2, "0");

  return `${y}-${m}-${day}`;
}

const REMINDER_WINDOW_DAYS = 15;

function reminderDaysLeft(client) {
  const remaining = daysUntil(client.expiredDate);

  return remaining <= REMINDER_WINDOW_DAYS ? remaining : null;
}

// ==================================================================
// Auth helpers
// ==================================================================

function signToken(user) {
  return jwt.sign(
    {
      sub: user._id.toString(),
      username: user.username,
    },
    JWT_SECRET,
    {
      expiresIn: TOKEN_EXPIRES_IN,
    },
  );
}

function requireAuth(req, res, next) {
  const header = req.headers.authorization || "";

  const [scheme, token] = header.split(" ");

  if (scheme !== "Bearer" || !token) {
    return res.status(401).json({
      message: "Authentication required. Please log in.",
    });
  }

  try {
    const payload = jwt.verify(token, JWT_SECRET);

    req.user = payload;

    next();
  } catch (e) {
    return res.status(401).json({
      message: "Session expired or invalid. Please log in again.",
    });
  }
}

async function ensureDefaultAdmin() {
  const count = await User.countDocuments();

  if (count > 0) return;

  const passwordHash = await bcrypt.hash(DEFAULT_ADMIN_PASSWORD, 10);

  await User.create({
    username: DEFAULT_ADMIN_USERNAME.toLowerCase(),

    passwordHash,
  });

  console.log(
    "=================================================================",
  );

  console.log(" No users found — a default admin account was created:");

  console.log(`   Username: ${DEFAULT_ADMIN_USERNAME}`);

  console.log(`   Password: ${DEFAULT_ADMIN_PASSWORD}`);

  console.log(" Please log in and change this password immediately.");

  console.log(
    "=================================================================",
  );
}

// ==================================================================
// Auth routes
// ==================================================================

app.get("/api/health", (_, res) =>
  res.json({
    ok: true,
  }),
);

app.post("/api/auth/login", async (req, res) => {
  try {
    const username = String(req.body.username || "")
      .trim()
      .toLowerCase();

    const password = String(req.body.password || "");

    if (!username || !password) {
      return res.status(400).json({
        message: "Username and password are required.",
      });
    }

    const user = await User.findOne({
      username,
    });

    if (!user) {
      return res.status(401).json({
        message: "Invalid username or password.",
      });
    }

    const ok = await bcrypt.compare(password, user.passwordHash);

    if (!ok) {
      return res.status(401).json({
        message: "Invalid username or password.",
      });
    }

    const token = signToken(user);

    res.json({
      token,
      username: user.username,
    });
  } catch (e) {
    res.status(500).json({
      message: e.message,
    });
  }
});

app.get("/api/auth/me", requireAuth, (req, res) => {
  res.json({
    username: req.user.username,
  });
});

app.post("/api/auth/change-password", requireAuth, async (req, res) => {
  try {
    const { currentPassword, newPassword } = req.body;

    if (!currentPassword || !newPassword || newPassword.length < 6) {
      return res.status(400).json({
        message:
          "Current password and a new password (min 6 characters) are required.",
      });
    }

    const user = await User.findById(req.user.sub);

    if (!user) {
      return res.status(404).json({
        message: "User not found.",
      });
    }

    const ok = await bcrypt.compare(currentPassword, user.passwordHash);

    if (!ok) {
      return res.status(401).json({
        message: "Current password is incorrect.",
      });
    }

    user.passwordHash = await bcrypt.hash(newPassword, 10);

    await user.save();

    res.json({
      message: "Password updated successfully.",
    });
  } catch (e) {
    res.status(500).json({
      message: e.message,
    });
  }
});

// ==================================================================
// Client management
// ==================================================================

app.get("/api/clients/options", requireAuth, (_, res) => {
  res.json(CLIENT_OPTIONS);
});

app.get("/api/clients", requireAuth, async (req, res) => {
  try {
    const search = String(req.query.search || "").trim();

    const filter = search
      ? {
          $or: [
            {
              partyName: new RegExp(search, "i"),
            },

            {
              firmName: new RegExp(search, "i"),
            },

            {
              mobileNo: new RegExp(search, "i"),
            },

            {
              emailId: new RegExp(search, "i"),
            },

            {
              licenseNumber: new RegExp(search, "i"),
            },

            {
              clientNumber: new RegExp(search, "i"),
            },

            {
              userId: new RegExp(search, "i"),
            },
          ],
        }
      : {};

    const clients = await Client.find(filter).sort({
      expiredDate: 1,
    });

    res.json(clients);
  } catch (e) {
    res.status(500).json({
      message: e.message,
    });
  }
});

app.post("/api/clients", requireAuth, async (req, res) => {
  try {
    const expiredDate = parseDate(req.body.expiredDate);

    if (!req.body.partyName || !String(req.body.partyName).trim()) {
      return res.status(400).json({
        message: "Party name is required.",
      });
    }

    if (!expiredDate) {
      return res.status(400).json({
        message: "A valid expired date is required.",
      });
    }

    const client = await Client.create({
      partyName: req.body.partyName,

      firmName: req.body.firmName,

      userId: req.body.userId,

      password: req.body.password,

      mobileNo: req.body.mobileNo,

      emailId: req.body.emailId,

      licenseNumber: req.body.licenseNumber,

      licenseType: req.body.licenseType,

      clientNumber: req.body.clientNumber,

      designation: req.body.designation,

      kob: req.body.kob,

      expiredDate,

      expiredDateFormat: normalizeDateFormat(req.body.expiredDateFormat),
    });

    res.status(201).json(client);
  } catch (e) {
    res.status(400).json({
      message: e.message,
    });
  }
});

app.put("/api/clients/:id", requireAuth, async (req, res) => {
  try {
    const existing = await Client.findById(req.params.id);

    if (!existing) {
      return res.status(404).json({
        message: "Client not found",
      });
    }

    const update = {
      partyName: req.body.partyName,

      firmName: req.body.firmName,

      userId: req.body.userId,

      password: req.body.password,

      mobileNo: req.body.mobileNo,

      emailId: req.body.emailId,

      licenseNumber: req.body.licenseNumber,

      licenseType: req.body.licenseType,

      clientNumber: req.body.clientNumber,

      designation: req.body.designation,

      kob: req.body.kob,

      updatedAt: new Date(),
    };

    if (Object.prototype.hasOwnProperty.call(req.body, "expiredDate")) {
      const expiredDate = parseDate(req.body.expiredDate);

      if (!expiredDate) {
        return res.status(400).json({
          message: "Valid expired date is required",
        });
      }

      update.expiredDate = expiredDate;

      update.expiredDateFormat = normalizeDateFormat(
        req.body.expiredDateFormat || existing.expiredDateFormat,
      );

      if (existing.expiredDate?.getTime() !== expiredDate.getTime()) {
        update.remindersSent = [];
      }
    }

    if (!update.partyName || !String(update.partyName).trim()) {
      return res.status(400).json({
        message: "Party name is required.",
      });
    }

    const client = await Client.findByIdAndUpdate(req.params.id, update, {
      new: true,
      runValidators: true,
    });

    res.json(client);
  } catch (e) {
    res.status(400).json({
      message: e.message,
    });
  }
});

app.delete("/api/clients/:id", requireAuth, async (req, res) => {
  try {
    const client = await Client.findByIdAndDelete(req.params.id);

    if (!client) {
      return res.status(404).json({
        message: "Client not found",
      });
    }

    res.json({
      message: "Client deleted.",
    });
  } catch (e) {
    res.status(400).json({
      message: e.message,
    });
  }
});

app.post("/api/clients/:id/renew", requireAuth, async (req, res) => {
  try {
    const existing = await Client.findById(req.params.id);

    if (!existing) {
      return res.status(404).json({
        message: "Client not found",
      });
    }

    const update = {
      renewed: true,

      lastRenewedAt: new Date(),

      dismissedOn: null,

      updatedAt: new Date(),
    };

    if (req.body.newExpiredDate) {
      const newExpiredDate = parseDate(req.body.newExpiredDate);

      if (!newExpiredDate) {
        return res.status(400).json({
          message: "Invalid new expired date.",
        });
      }

      update.expiredDate = newExpiredDate;

      update.expiredDateFormat = normalizeDateFormat(
        existing.expiredDateFormat,
      );

      if (existing.expiredDate?.getTime() !== newExpiredDate.getTime()) {
        update.remindersSent = [];
      }
    }

    const client = await Client.findByIdAndUpdate(req.params.id, update, {
      new: true,
    });

    res.json(client);
  } catch (e) {
    res.status(400).json({
      message: e.message,
    });
  }
});

// ==================================================================
// Reminder feed
// ==================================================================

app.post("/api/clients/:id/dismiss-today", async (req, res) => {
  try {
    const client = await Client.findByIdAndUpdate(
      req.params.id,
      {
        dismissedOn: dateKey(),

        updatedAt: new Date(),
      },
      {
        new: true,
      },
    );

    res.json(client);
  } catch (e) {
    res.status(400).json({
      message: e.message,
    });
  }
});

app.get("/api/reminders", async (_, res) => {
  try {
    const clients = await Client.find({
      expiredDate: {
        $exists: true,
      },
    });

    const reminders = clients

      .map((client) => {
        const remaining = reminderDaysLeft(client);

        return {
          ...client.toObject(),

          daysRemaining: remaining,

          dismissedToday: client.dismissedOn === dateKey(),
        };
      })

      .filter(
        (client) => client.daysRemaining !== null && !client.dismissedToday,
      );

    res.json(reminders);
  } catch (e) {
    res.status(500).json({
      message: e.message,
    });
  }
});

// ==================================================================
// Export Excel
// ==================================================================

const getHyphenatedDateTime = (ts = Date.now()) => {
  const date = new Date(ts);

  const d = date
    .toLocaleDateString("en-GB", {
      day: "2-digit",
      month: "short",
      year: "numeric",
    })
    .replace(/ /g, "-"); // "21-Sep-2026"

  const t = date
    .toLocaleTimeString("en-US", {
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hour12: true,
    })
    .toLowerCase() // "01:45:12 pm"
    .replace(/:/g, "-") // "01-45-12 pm"
    .replace(/\s+/g, "-"); // "01-45-12-pm"

  return `${d}-${t}`;
};

// Output: "21-Sep-2026-01-45-12-pm"

app.get("/api/export/excel", requireAuth, async (_, res) => {
  try {
    const clients = await Client.find({})
      .sort({
        expiredDate: 1,
      })
      .lean();

    const headers = [
      "party_name",
      "firm_name",
      "user_id",
      "password",
      "mobile_no",
      "email_id",
      "license_number",
      "license_type",
      "client_number",
      "designation",
      "kob",
      "expired_date",
      "last_renewed_at",
    ];

    const data = clients.map((c) => ({
      party_name: c.partyName ?? "",

      firm_name: c.firmName ?? "",

      user_id: c.userId ?? "",

      password: c.password ?? "",

      mobile_no: c.mobileNo ?? "",

      email_id: c.emailId ?? "",

      license_number: c.licenseNumber ?? "",

      license_type: c.licenseType ?? "",

      client_number: c.clientNumber ?? "",

      designation: c.designation ?? "",

      kob: c.kob ?? "",

      expired_date: c.expiredDate ? new Date(c.expiredDate) : "",

      last_renewed_at: c.lastRenewedAt ? new Date(c.lastRenewedAt) : "",
    }));

    const ws = XLSX.utils.json_to_sheet(data, {
      header: headers,
      cellDates: true,
    });

    clients.forEach((c, i) => {
      const row = i + 2;

      if (c.expiredDate && ws[`L${row}`]) {
        ws[`L${row}`].z = normalizeDateFormat(c.expiredDateFormat);
      }

      if (c.lastRenewedAt && ws[`M${row}`]) {
        ws[`M${row}`].z = "dd-mm-yyyy hh:mm";
      }
    });

    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "Clients");

    const buffer = XLSX.write(wb, {
      type: "buffer",
      bookType: "xlsx",
      cellDates: true,
    });

    const filename = `clients-${getHyphenatedDateTime()}.xlsx`;
    // → "clients-21-Sep-2026-01-45-12-pm.xlsx"

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );

    // RFC 5987 encoding for safe non-ASCII / special chars
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${filename}"; filename*=UTF-8''${encodeURIComponent(filename)}`,
    );

    res.send(buffer);
  } catch (e) {
    res.status(500).json({ message: e.message });
  }
});

// ==================================================================
// Email reminders - Nodemailer
// ==================================================================

const EMAIL_REMINDER_DAYS = [15, 10, 5, 3, 1];

function currentMilestone(remaining) {
  const sorted = [...EMAIL_REMINDER_DAYS].sort((a, b) => a - b);

  if (remaining < 0) {
    return null;
  }

  if (remaining > sorted[sorted.length - 1]) {
    return null;
  }

  for (const m of sorted) {
    if (remaining <= m) {
      return m;
    }
  }

  return null;
}

const EMPLOYEE_NAME = process.env.EMPLOYEE_NAME || "Team";

const EMPLOYEE_EMAIL =
  process.env.EMPLOYEE_EMAIL || process.env.SUJIT_EMAIL || "";

if (!EMPLOYEE_EMAIL) {
  console.warn(
    "[mailer] EMPLOYEE_EMAIL (or SUJIT_EMAIL) is not set in .env — " +
      "the employee will NOT receive renewal-reminder emails.",
  );
}

// ==================================================================
// Nodemailer SMTP
// ==================================================================

const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,

  port: Number(process.env.SMTP_PORT) || 587,

  secure: process.env.SMTP_SECURE === "true",

  auth: {
    user: process.env.SMTP_USER,

    pass: process.env.SMTP_PASS,
  },

  tls: {
    rejectUnauthorized: false,
  },

  connectionTimeout: 15000,

  greetingTimeout: 15000,

  socketTimeout: 20000,
});

transporter.verify((err) => {
  if (err) {
    console.error(
      "[mailer] SMTP connection FAILED — reminder emails will not send:",
      err.message,
    );
  } else {
    console.log("[mailer] SMTP connection verified OK");
  }
});

// ==================================================================
// Email formatting
// ==================================================================

function formatDate(date, pattern = DEFAULT_DATE_FORMAT) {
  return date ? formatDateByPattern(date, pattern) : "N/A";
}

function buildReminderEmail(client, daysRemaining) {
  const dayWord = daysRemaining === 1 ? "day" : "days";

  const subject = `Reminder: License for ${client.partyName} expires in ${daysRemaining} ${dayWord}`;

  const text =
    `Hi ${client.partyName},\n\n` +
    `This is a reminder that the license/registration for ` +
    `${client.firmName || client.partyName} ` +
    `is set to expire on ` +
    `${formatDate(client.expiredDate, client.expiredDateFormat)} ` +
    `(${daysRemaining} ${dayWord} from today).\n\n` +
    `Please renew soon to avoid any interruption.\n\n` +
    `Thank you.`;

  const html = `

    <div
      style="
        font-family: Arial, sans-serif;
        font-size: 15px;
        color: #222;
      "
    >

      <p>
        Hi ${client.partyName},
      </p>

      <p>

        This is a reminder that the
        license/registration for

        <strong>
          ${client.firmName || client.partyName}
        </strong>

        is set to expire on

        <strong>
          ${formatDate(client.expiredDate, client.expiredDateFormat)}
        </strong>

        (<strong>
          ${daysRemaining} ${dayWord}
        </strong>
        from today).

      </p>

      <p>
        Please renew soon to avoid any interruption in service.
      </p>

      <p>
        Thank you.
      </p>

    </div>

  `;

  return {
    subject,
    text,
    html,
  };
}

function buildEmployeeReminderEmail(client, daysRemaining, employeeName) {
  const dayWord = daysRemaining === 1 ? "day" : "days";

  const subject = `Renewal Alert: ${client.partyName} expires in ${daysRemaining} ${dayWord}`;

  const text =
    `Hii ${employeeName},\n\n` +
    `The party name is ${client.partyName}, ` +
    `their plan is expiring in just ` +
    `${daysRemaining} ${dayWord}.\n\n` +
    `Firm Name: ${client.firmName || "-"}\n` +
    `Mobile No.: ${client.mobileNo || "-"}\n` +
    `Client Email: ${client.emailId || "-"}\n` +
    `License Number: ${client.licenseNumber || "-"}\n` +
    `Expired Date: ${formatDate(
      client.expiredDate,
      client.expiredDateFormat,
    )}\n\n` +
    `Please follow up on the renewal.`;

  const html = `

    <div
      style="
        font-family: Arial, sans-serif;
        font-size: 15px;
        color: #222;
      "
    >

      <p>
        Hii ${employeeName},
      </p>

      <p>

        The party name is
        <strong>
          ${client.partyName}
        </strong>,

        their plan is expiring in just

        <strong>
          ${daysRemaining} ${dayWord}
        </strong>.

      </p>


      <table
        style="
          border-collapse: collapse;
          margin-top: 10px;
        "
      >

        <tr>
          <td
            style="
              padding:4px 10px 4px 0;
              color:#667085;
            "
          >
            Firm Name
          </td>

          <td>
            ${client.firmName || "-"}
          </td>
        </tr>


        <tr>

          <td
            style="
              padding:4px 10px 4px 0;
              color:#667085;
            "
          >
            Mobile No.
          </td>

          <td>
            ${client.mobileNo || "-"}
          </td>

        </tr>


        <tr>

          <td
            style="
              padding:4px 10px 4px 0;
              color:#667085;
            "
          >
            Client Email
          </td>

          <td>
            ${client.emailId || "-"}
          </td>

        </tr>


        <tr>

          <td
            style="
              padding:4px 10px 4px 0;
              color:#667085;
            "
          >
            License Number
          </td>

          <td>
            ${client.licenseNumber || "-"}
          </td>

        </tr>


        <tr>

          <td
            style="
              padding:4px 10px 4px 0;
              color:#667085;
            "
          >
            Expired Date
          </td>

          <td>
            ${formatDate(client.expiredDate, client.expiredDateFormat)}
          </td>

        </tr>

      </table>


      <p
        style="
          margin-top: 14px;
        "
      >
        Please follow up on the renewal.
      </p>

    </div>

  `;

  return {
    subject,
    text,
    html,
  };
}

async function sendReminderEmailToClient(client, daysRemaining) {
  const { subject, text, html } = buildReminderEmail(client, daysRemaining);

  await transporter.sendMail({
    from: process.env.MAIL_FROM || process.env.SMTP_USER,

    to: client.emailId,

    subject,

    text,

    html,
  });
}

async function sendReminderEmailToEmployee(client, daysRemaining) {
  const { subject, text, html } = buildEmployeeReminderEmail(
    client,
    daysRemaining,
    EMPLOYEE_NAME,
  );

  await transporter.sendMail({
    from: process.env.MAIL_FROM || process.env.SMTP_USER,

    to: EMPLOYEE_EMAIL,

    subject,

    text,

    html,
  });
}

// ==================================================================
// Core reminder job
// ==================================================================

async function runReminderEmailJob() {
  const results = {
    sent: [],

    skipped: [],

    failed: [],

    employeeNotified: [],

    employeeFailed: [],
  };

  const clients = await Client.find({
    expiredDate: {
      $exists: true,
    },
  });

  for (const client of clients) {
    const remaining = daysUntil(client.expiredDate);

    const alreadySent = client.remindersSent || [];

    const milestone = currentMilestone(remaining);

    if (milestone === null) {
      continue;
    }

    if (alreadySent.includes(milestone)) {
      continue;
    }

    // --------------------------------------------------------------
    // Client email
    // --------------------------------------------------------------

    if (!client.emailId) {
      results.skipped.push({
        partyName: client.partyName,

        milestone,

        reason: "no email on file",
      });
    } else {
      try {
        await sendReminderEmailToClient(client, milestone);

        results.sent.push({
          partyName: client.partyName,

          emailId: client.emailId,

          daysRemaining: milestone,
        });
      } catch (err) {
        results.failed.push({
          partyName: client.partyName,

          emailId: client.emailId,

          milestone,

          error: err.message,
        });
      }
    }

    // --------------------------------------------------------------
    // Employee email
    // --------------------------------------------------------------

    if (EMPLOYEE_EMAIL) {
      try {
        await sendReminderEmailToEmployee(client, milestone);

        results.employeeNotified.push({
          partyName: client.partyName,

          employeeEmail: EMPLOYEE_EMAIL,

          daysRemaining: milestone,
        });
      } catch (err) {
        results.employeeFailed.push({
          partyName: client.partyName,

          employeeEmail: EMPLOYEE_EMAIL,

          milestone,

          error: err.message,
        });
      }
    }

    // --------------------------------------------------------------
    // Mark milestone as processed
    // --------------------------------------------------------------

    client.remindersSent = [...alreadySent, milestone];

    await client.save();
  }

  return results;
}

// ==================================================================
// Manual email reminder endpoint
// ==================================================================

app.post("/api/reminders/send-emails", requireAuth, async (req, res) => {
  try {
    if (req.query.force === "1" || req.query.force === "true") {
      await Client.updateMany(
        {},
        {
          $set: {
            remindersSent: [],
          },
        },
      );

      console.log("[reminders] force=1 — cleared remindersSent on all clients");
    }

    const results = await runReminderEmailJob();

    res.json(results);
  } catch (e) {
    res.status(500).json({
      message: e.message,
    });
  }
});

// ==================================================================
// Reminder debug endpoint
// ==================================================================

app.get("/api/reminders/debug", requireAuth, async (_, res) => {
  try {
    const clients = await Client.find({
      expiredDate: {
        $exists: true,
      },
    }).lean();

    const now = new Date();

    const rows = clients.map((c) => {
      const remaining = daysUntil(c.expiredDate);

      const alreadySent = c.remindersSent || [];

      const milestone = currentMilestone(remaining);

      const willSend = milestone !== null && !alreadySent.includes(milestone);

      return {
        partyName: c.partyName,

        emailId: c.emailId,

        expiredDateISO: c.expiredDate,

        expiredDateLocal: new Date(c.expiredDate).toString(),

        todayLocal: now.toString(),

        daysRemaining: remaining,

        currentMilestone: milestone,

        remindersSent: alreadySent,

        wouldSendMilestone: willSend ? milestone : null,

        wouldEmailClient: Boolean(willSend && c.emailId),

        wouldEmailEmployee: Boolean(willSend && EMPLOYEE_EMAIL),
      };
    });

    res.json({
      employeeEmail: EMPLOYEE_EMAIL || null,

      reminderDays: EMAIL_REMINDER_DAYS,

      rows,
    });
  } catch (e) {
    res.status(500).json({
      message: e.message,
    });
  }
});

// ==================================================================
// EXTERNAL CRON - cron-job.org
// ==================================================================
//
// cron-job.org will call:
//
// GET
// /api/reminders/cron?key=YOUR_SECRET
//
// Example:
//
// https://service-renewal-reminder-backend.onrender.com/api/reminders/cron?key=YOUR_SECRET
// ==================================================================

app.get("/api/reminders/cron", async (req, res) => {
  try {
    const providedKey = String(req.query.key || "");

    const expectedKey = String(process.env.REMINDER_CRON_KEY || "");

    // ------------------------------------------------------------
    // Make sure the Render environment variable exists
    // ------------------------------------------------------------

    if (!expectedKey) {
      console.error("[external-cron] REMINDER_CRON_KEY is not configured.");

      return res.status(500).json({
        success: false,

        message: "REMINDER_CRON_KEY is not configured on the server.",
      });
    }

    // ------------------------------------------------------------
    // Validate cron secret
    // ------------------------------------------------------------

    if (!providedKey || providedKey !== expectedKey) {
      console.warn("[external-cron] Unauthorized cron request.");

      return res.status(401).json({
        success: false,

        message: "Unauthorized.",
      });
    }

    console.log(
      "===============================================================",
    );

    console.log(
      `[external-cron] Reminder job started: ${new Date().toISOString()}`,
    );

    console.log(
      "===============================================================",
    );

    // ------------------------------------------------------------
    // Run reminder email job
    // ------------------------------------------------------------

    const results = await runReminderEmailJob();

    // ------------------------------------------------------------
    // Log results
    // ------------------------------------------------------------

    console.log(`[external-cron] Client emails sent: ${results.sent.length}`);

    console.log(
      `[external-cron] Client emails skipped: ${results.skipped.length}`,
    );

    console.log(
      `[external-cron] Client emails failed: ${results.failed.length}`,
    );

    console.log(
      `[external-cron] Employee emails sent: ${results.employeeNotified.length}`,
    );

    console.log(
      `[external-cron] Employee emails failed: ${results.employeeFailed.length}`,
    );

    results.sent.forEach((item) => {
      console.log(
        `[external-cron] CLIENT -> ${item.partyName} <${item.emailId}> (${item.daysRemaining}d)`,
      );
    });

    results.failed.forEach((item) => {
      console.error(
        `[external-cron] CLIENT FAILED -> ${item.partyName} <${item.emailId}> (${item.milestone}d): ${item.error}`,
      );
    });

    results.skipped.forEach((item) => {
      console.warn(
        `[external-cron] SKIPPED -> ${item.partyName} (${item.milestone}d): ${item.reason}`,
      );
    });

    results.employeeNotified.forEach((item) => {
      console.log(
        `[external-cron] EMPLOYEE -> ${EMPLOYEE_NAME} <${item.employeeEmail}> about ${item.partyName} (${item.daysRemaining}d)`,
      );
    });

    results.employeeFailed.forEach((item) => {
      console.error(
        `[external-cron] EMPLOYEE FAILED -> <${item.employeeEmail}> about ${item.partyName} (${item.milestone}d): ${item.error}`,
      );
    });

    console.log(
      `[external-cron] Reminder job completed: ${new Date().toISOString()}`,
    );

    return res.json({
      success: true,

      message: "Reminder email job completed.",

      executedAt: new Date().toISOString(),

      summary: {
        clientEmailsSent: results.sent.length,

        clientEmailsSkipped: results.skipped.length,

        clientEmailsFailed: results.failed.length,

        employeeEmailsSent: results.employeeNotified.length,

        employeeEmailsFailed: results.employeeFailed.length,
      },

      results,
    });
  } catch (error) {
    console.error("[external-cron] Reminder job failed:", error);

    return res.status(500).json({
      success: false,

      message: error.message || "Reminder job failed.",
    });
  }
});

// ==================================================================
// IMPORTANT
// ==================================================================
// INTERNAL node-cron has been intentionally REMOVED.
// cron-job.org will trigger:
// GET /api/reminders/cron?key=YOUR_SECRET
// This prevents duplicate emails when Render is running.
// ==================================================================
// ==================================================================
// Startup
// ==================================================================

mongoose
  .connect(process.env.MONGO_URI)

  .then(async () => {
    await ensureDefaultAdmin();

    app.listen(port, "0.0.0.0", () => {
      console.log(`Backend running on port ${port}`);

      console.log(`External cron endpoint: /api/reminders/cron`);

      console.log(`SMTP host: ${process.env.SMTP_HOST || "not configured"}`);
    });
  })

  .catch((err) => {
    console.error("MongoDB connection failed:", err.message);

    process.exit(1);
  });


// take as it is just add here cron-org for external cron

// import express from "express";
// import cors from "cors";
// import dotenv from "dotenv";
// import mongoose from "mongoose";
// import XLSX from "xlsx";
// import nodemailer from "nodemailer";
// import cron from "node-cron";
// import bcrypt from "bcryptjs";
// import jwt from "jsonwebtoken";

// dotenv.config();

// const app = express();
// const port = process.env.PORT || 5000;

// app.use(cors({ origin: process.env.CORS_ORIGIN?.split(",") || "*" }));
// app.use(express.json());

// // ==================================================================
// // Auth config
// // ==================================================================
// const JWT_SECRET = process.env.JWT_SECRET || "please-change-this-secret-in-env";
// const TOKEN_EXPIRES_IN = process.env.TOKEN_EXPIRES_IN || "12h";
// const DEFAULT_ADMIN_USERNAME = process.env.ADMIN_USERNAME || "admin";
// const DEFAULT_ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "Admin@123";

// if (!process.env.JWT_SECRET) {
//   console.warn(
//     "[auth] WARNING: JWT_SECRET is not set in .env — using an insecure default. " +
//     "Set JWT_SECRET in backend/.env before using this in production."
//   );
// }

// // ==================================================================
// // Schemas
// // ==================================================================
// const userSchema = new mongoose.Schema({
//   username: { type: String, required: true, unique: true, trim: true, lowercase: true },
//   passwordHash: { type: String, required: true },
//   createdAt: { type: Date, default: Date.now }
// });
// const User = mongoose.model("User", userSchema);

// const clientSchema = new mongoose.Schema({
//   partyName: { type: String, required: true, trim: true },
//   firmName: { type: String, trim: true, default: "" },
//   userId: { type: String, trim: true, default: "" },
//   password: { type: String, default: "" },
//   mobileNo: { type: String, trim: true, default: "" },
//   emailId: { type: String, trim: true, default: "" },
//   licenseNumber: { type: String, trim: true, default: "" },
//   licenseType: { type: String, trim: true, default: "" },
//   clientNumber: { type: String, trim: true, default: "" },
//   designation: { type: String, trim: true, default: "" },
//   kob: { type: String, trim: true, default: "" },

//   expiredDate: { type: Date, required: true },
//   expiredDateFormat: { type: String, default: "dd-mm-yyyy" },

//   lastRenewedAt: { type: Date, default: null },

//   renewed: { type: Boolean, default: false },
//   dismissedOn: { type: String, default: null },
//   remindersSent: { type: [Number], default: [] },

//   createdAt: { type: Date, default: Date.now },
//   updatedAt: { type: Date, default: Date.now }
// });

// const Client = mongoose.model("Client", clientSchema);

// const DEFAULT_DATE_FORMAT = "dd-mm-yyyy";

// // ==================================================================
// // Date helpers — (expiryDate − todayDate) in whole calendar days.
// // ==================================================================
// //   today  = new Date(Y, M, D, 0, 0, 0)   // local midnight
// //   expiry = new Date(Y, M, D, 0, 0, 0)   // local midnight
// //   days   = Math.round((expiry - today) / 86400000)
// //
// // Example: today = 16 Sep 2026, expiry = 21 Sep 2026 → 5 days.
// // ==================================================================

// function atMidnight(d) {
//   const x = d instanceof Date ? d : new Date(d);
//   return new Date(x.getFullYear(), x.getMonth(), x.getDate(), 0, 0, 0, 0);
// }

// function daysUntil(date) {
//   if (!date) return null;
//   const today = atMidnight(new Date());
//   const expiry = atMidnight(date);
//   return Math.round((expiry - today) / 86400000);
// }

// function isValidDateParts(year, month, day) {
//   if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return false;
//   const d = new Date(year, month - 1, day);
//   return d.getFullYear() === year && d.getMonth() === month - 1 && d.getDate() === day;
// }

// function parseDate(value) {
//   if (value === null || value === undefined || value === "") return null;

//   if (value instanceof Date) {
//     if (Number.isNaN(value.getTime())) return null;
//     return new Date(value.getFullYear(), value.getMonth(), value.getDate(), 0, 0, 0, 0);
//   }

//   if (typeof value === "number") {
//     const parsed = XLSX.SSF.parse_date_code(value);
//     if (!parsed || !isValidDateParts(parsed.y, parsed.m, parsed.d)) return null;
//     return new Date(parsed.y, parsed.m - 1, parsed.d, 0, 0, 0, 0);
//   }

//   const text = String(value).trim();

//   let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
//   if (match) {
//     const [, y, m, d] = match.map(Number);
//     return isValidDateParts(y, m, d) ? new Date(y, m - 1, d, 0, 0, 0, 0) : null;
//   }

//   match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
//   if (match) {
//     let [, d, m, y] = match;
//     if (y.length === 2) y = `20${y}`;
//     d = Number(d); m = Number(m); y = Number(y);
//     return isValidDateParts(y, m, d) ? new Date(y, m - 1, d, 0, 0, 0, 0) : null;
//   }

//   const date = new Date(text);
//   return Number.isNaN(date.getTime())
//     ? null
//     : new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
// }

// function normalizeDateFormat(format) {
//   if (!format || typeof format !== "string") return DEFAULT_DATE_FORMAT;
//   const cleaned = format.trim().replace(/\\\-/g, "-").replace(/\\\//g, "/");
//   const lower = cleaned.toLowerCase();
//   if (lower.includes("yy") && (lower.includes("dd") || lower.includes("d")) &&
//       (lower.includes("mm") || lower.includes("m"))) {
//     return cleaned;
//   }
//   return DEFAULT_DATE_FORMAT;
// }

// function formatDateByPattern(value, pattern = DEFAULT_DATE_FORMAT) {
//   const date = parseDate(value);
//   if (!date) return "";
//   const d = String(date.getDate()).padStart(2, "0");
//   const m = String(date.getMonth() + 1).padStart(2, "0");
//   const y = String(date.getFullYear());
//   const yy = y.slice(-2);

//   let out = pattern;
//   out = out.replace(/yyyy/gi, y).replace(/yy/gi, yy);
//   out = out.replace(/dd/gi, d);
//   out = out.replace(/mm/gi, m);
//   out = out.replace(/d/g, String(date.getDate()));
//   out = out.replace(/m/g, String(date.getMonth() + 1));
//   return out;
// }

// function dateKey(date = new Date()) {
//   const d = date instanceof Date ? date : new Date(date);
//   const y = d.getFullYear();
//   const m = String(d.getMonth() + 1).padStart(2, "0");
//   const day = String(d.getDate()).padStart(2, "0");
//   return `${y}-${m}-${day}`;
// }

// const REMINDER_WINDOW_DAYS = 15;
// function reminderDaysLeft(client) {
//   const remaining = daysUntil(client.expiredDate);
//   return remaining <= REMINDER_WINDOW_DAYS ? remaining : null;
// }

// // ==================================================================
// // Auth helpers & middleware
// // ==================================================================
// function signToken(user) {
//   return jwt.sign({ sub: user._id.toString(), username: user.username }, JWT_SECRET, {
//     expiresIn: TOKEN_EXPIRES_IN
//   });
// }

// function requireAuth(req, res, next) {
//   const header = req.headers.authorization || "";
//   const [scheme, token] = header.split(" ");
//   if (scheme !== "Bearer" || !token) {
//     return res.status(401).json({ message: "Authentication required. Please log in." });
//   }
//   try {
//     const payload = jwt.verify(token, JWT_SECRET);
//     req.user = payload;
//     next();
//   } catch (e) {
//     return res.status(401).json({ message: "Session expired or invalid. Please log in again." });
//   }
// }

// async function ensureDefaultAdmin() {
//   const count = await User.countDocuments();
//   if (count > 0) return;
//   const passwordHash = await bcrypt.hash(DEFAULT_ADMIN_PASSWORD, 10);
//   await User.create({ username: DEFAULT_ADMIN_USERNAME.toLowerCase(), passwordHash });
//   console.log("=================================================================");
//   console.log(" No users found — a default admin account was created:");
//   console.log(`   Username: ${DEFAULT_ADMIN_USERNAME}`);
//   console.log(`   Password: ${DEFAULT_ADMIN_PASSWORD}`);
//   console.log(" Please log in and change this password immediately.");
//   console.log("=================================================================");
// }

// // ==================================================================
// // Auth routes (public)
// // ==================================================================
// app.get("/api/health", (_, res) => res.json({ ok: true }));

// app.post("/api/auth/login", async (req, res) => {
//   try {
//     const username = String(req.body.username || "").trim().toLowerCase();
//     const password = String(req.body.password || "");
//     if (!username || !password) {
//       return res.status(400).json({ message: "Username and password are required." });
//     }

//     const user = await User.findOne({ username });
//     if (!user) return res.status(401).json({ message: "Invalid username or password." });

//     const ok = await bcrypt.compare(password, user.passwordHash);
//     if (!ok) return res.status(401).json({ message: "Invalid username or password." });

//     const token = signToken(user);
//     res.json({ token, username: user.username });
//   } catch (e) {
//     res.status(500).json({ message: e.message });
//   }
// });

// app.get("/api/auth/me", requireAuth, (req, res) => {
//   res.json({ username: req.user.username });
// });

// app.post("/api/auth/change-password", requireAuth, async (req, res) => {
//   try {
//     const { currentPassword, newPassword } = req.body;
//     if (!currentPassword || !newPassword || newPassword.length < 6) {
//       return res.status(400).json({
//         message: "Current password and a new password (min 6 characters) are required."
//       });
//     }
//     const user = await User.findById(req.user.sub);
//     if (!user) return res.status(404).json({ message: "User not found." });

//     const ok = await bcrypt.compare(currentPassword, user.passwordHash);
//     if (!ok) return res.status(401).json({ message: "Current password is incorrect." });

//     user.passwordHash = await bcrypt.hash(newPassword, 10);
//     await user.save();
//     res.json({ message: "Password updated successfully." });
//   } catch (e) {
//     res.status(500).json({ message: e.message });
//   }
// });

// // ==================================================================
// // Client management routes (protected)
// // ==================================================================
// app.get("/api/clients", requireAuth, async (req, res) => {
//   try {
//     const search = String(req.query.search || "").trim();
//     const filter = search
//       ? { $or: [
//           { partyName: new RegExp(search, "i") },
//           { firmName: new RegExp(search, "i") },
//           { mobileNo: new RegExp(search, "i") },
//           { emailId: new RegExp(search, "i") },
//           { licenseNumber: new RegExp(search, "i") },
//           { clientNumber: new RegExp(search, "i") }
//         ] }
//       : {};
//     const clients = await Client.find(filter).sort({ expiredDate: 1 });
//     res.json(clients);
//   } catch (e) {
//     res.status(500).json({ message: e.message });
//   }
// });

// app.post("/api/clients", requireAuth, async (req, res) => {
//   try {
//     const expiredDate = parseDate(req.body.expiredDate);
//     if (!req.body.partyName || !String(req.body.partyName).trim()) {
//       return res.status(400).json({ message: "Party name is required." });
//     }
//     if (!expiredDate) {
//       return res.status(400).json({ message: "A valid expired date is required." });
//     }

//     const client = await Client.create({
//       partyName: req.body.partyName,
//       firmName: req.body.firmName,
//       userId: req.body.userId,
//       password: req.body.password,
//       mobileNo: req.body.mobileNo,
//       emailId: req.body.emailId,
//       licenseNumber: req.body.licenseNumber,
//       licenseType: req.body.licenseType,
//       clientNumber: req.body.clientNumber,
//       designation: req.body.designation,
//       kob: req.body.kob,
//       expiredDate,
//       expiredDateFormat: normalizeDateFormat(req.body.expiredDateFormat)
//     });
//     res.status(201).json(client);
//   } catch (e) {
//     res.status(400).json({ message: e.message });
//   }
// });

// app.put("/api/clients/:id", requireAuth, async (req, res) => {
//   try {
//     const existing = await Client.findById(req.params.id);
//     if (!existing) return res.status(404).json({ message: "Client not found" });

//     const update = {
//       partyName: req.body.partyName,
//       firmName: req.body.firmName,
//       userId: req.body.userId,
//       password: req.body.password,
//       mobileNo: req.body.mobileNo,
//       emailId: req.body.emailId,
//       licenseNumber: req.body.licenseNumber,
//       licenseType: req.body.licenseType,
//       clientNumber: req.body.clientNumber,
//       designation: req.body.designation,
//       kob: req.body.kob,
//       updatedAt: new Date()
//     };

//     if (Object.prototype.hasOwnProperty.call(req.body, "expiredDate")) {
//       const expiredDate = parseDate(req.body.expiredDate);
//       if (!expiredDate) return res.status(400).json({ message: "Valid expired date is required" });
//       update.expiredDate = expiredDate;
//       update.expiredDateFormat = normalizeDateFormat(
//         req.body.expiredDateFormat || existing.expiredDateFormat
//       );
//       if (existing.expiredDate?.getTime() !== expiredDate.getTime()) {
//         update.remindersSent = [];
//       }
//     }

//     if (!update.partyName || !String(update.partyName).trim()) {
//       return res.status(400).json({ message: "Party name is required." });
//     }

//     const client = await Client.findByIdAndUpdate(req.params.id, update, { new: true, runValidators: true });
//     res.json(client);
//   } catch (e) {
//     res.status(400).json({ message: e.message });
//   }
// });

// app.delete("/api/clients/:id", requireAuth, async (req, res) => {
//   try {
//     const client = await Client.findByIdAndDelete(req.params.id);
//     if (!client) return res.status(404).json({ message: "Client not found" });
//     res.json({ message: "Client deleted." });
//   } catch (e) {
//     res.status(400).json({ message: e.message });
//   }
// });

// app.post("/api/clients/:id/renew", requireAuth, async (req, res) => {
//   try {
//     const existing = await Client.findById(req.params.id);
//     if (!existing) return res.status(404).json({ message: "Client not found" });

//     const update = {
//       renewed: true,
//       lastRenewedAt: new Date(),
//       dismissedOn: null,
//       updatedAt: new Date()
//     };

//     if (req.body.newExpiredDate) {
//       const newExpiredDate = parseDate(req.body.newExpiredDate);
//       if (!newExpiredDate) return res.status(400).json({ message: "Invalid new expired date." });
//       update.expiredDate = newExpiredDate;
//       update.expiredDateFormat = normalizeDateFormat(existing.expiredDateFormat);
//       if (existing.expiredDate?.getTime() !== newExpiredDate.getTime()) {
//         update.remindersSent = [];
//       }
//     }

//     const client = await Client.findByIdAndUpdate(req.params.id, update, { new: true });
//     res.json(client);
//   } catch (e) {
//     res.status(400).json({ message: e.message });
//   }
// });

// // ==================================================================
// // Reminder feed used by the Chrome extension popup.
// // ==================================================================
// app.post("/api/clients/:id/dismiss-today", async (req, res) => {
//   try {
//     const client = await Client.findByIdAndUpdate(
//       req.params.id,
//       { dismissedOn: dateKey(), updatedAt: new Date() },
//       { new: true }
//     );
//     res.json(client);
//   } catch (e) {
//     res.status(400).json({ message: e.message });
//   }
// });

// app.get("/api/reminders", async (_, res) => {
//   try {
//     const clients = await Client.find({ expiredDate: { $exists: true } });
//     const reminders = clients
//       .map(client => {
//         const remaining = reminderDaysLeft(client);
//         return {
//           ...client.toObject(),
//           daysRemaining: remaining,
//           dismissedToday: client.dismissedOn === dateKey()
//         };
//       })
//       .filter(client => client.daysRemaining !== null && !client.dismissedToday);
//     res.json(reminders);
//   } catch (e) {
//     res.status(500).json({ message: e.message });
//   }
// });

// // ==================================================================
// // Export to Excel (protected)
// // ==================================================================
// app.get("/api/export/excel", requireAuth, async (_, res) => {
//   try {
//     const clients = await Client.find({}).sort({ expiredDate: 1 }).lean();
//     const headers = [
//       "party_name", "firm_name", "user_id", "password", "mobile_no", "email_id",
//       "license_number", "license_type", "client_number", "designation", "kob",
//       "expired_date", "last_renewed_at"
//     ];
//     const data = clients.map(c => ({
//       party_name: c.partyName ?? "",
//       firm_name: c.firmName ?? "",
//       user_id: c.userId ?? "",
//       password: c.password ?? "",
//       mobile_no: c.mobileNo ?? "",
//       email_id: c.emailId ?? "",
//       license_number: c.licenseNumber ?? "",
//       license_type: c.licenseType ?? "",
//       client_number: c.clientNumber ?? "",
//       designation: c.designation ?? "",
//       kob: c.kob ?? "",
//       expired_date: c.expiredDate ? new Date(c.expiredDate) : "",
//       last_renewed_at: c.lastRenewedAt ? new Date(c.lastRenewedAt) : ""
//     }));

//     const ws = XLSX.utils.json_to_sheet(data, { header: headers, cellDates: true });
//     clients.forEach((c, i) => {
//       const row = i + 2;
//       if (c.expiredDate && ws[`L${row}`]) {
//         ws[`L${row}`].z = normalizeDateFormat(c.expiredDateFormat);
//       }
//       if (c.lastRenewedAt && ws[`M${row}`]) {
//         ws[`M${row}`].z = "dd-mm-yyyy hh:mm";
//       }
//     });

//     const wb = XLSX.utils.book_new();
//     XLSX.utils.book_append_sheet(wb, ws, "Clients");
//     const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx", cellDates: true });
//     res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//     res.setHeader("Content-Disposition", 'attachment; filename="clients.xlsx"');
//     res.send(buffer);
//   } catch (e) {
//     res.status(500).json({ message: e.message });
//   }
// });

// // mail service start here

// // ==================================================================
// // Email reminders (Nodemailer)
// // ==================================================================
// // Milestones, in descending order of "how far out" they represent.
// const EMAIL_REMINDER_DAYS = [15, 10, 5];

// // Return the SINGLE milestone the client is currently inside, or null
// // if none applies. Windows:
// //   remaining > 15      → null   (nothing yet)
// //   11 ≤ remaining ≤ 15 → 15
// //    6 ≤ remaining ≤ 10 → 10
// //    0 ≤ remaining ≤  5 → 5
// //   remaining < 0       → null   (already expired — no email)
// function currentMilestone(remaining) {
//   const sorted = [...EMAIL_REMINDER_DAYS].sort((a, b) => a - b); // [5, 10, 15]
//   if (remaining < 0) return null;
//   if (remaining > sorted[sorted.length - 1]) return null; // above 15
//   for (const m of sorted) {
//     if (remaining <= m) return m;
//   }
//   return null;
// }

// const EMPLOYEE_NAME = process.env.EMPLOYEE_NAME || "Team";
// const EMPLOYEE_EMAIL = process.env.EMPLOYEE_EMAIL || process.env.SUJIT_EMAIL || "";

// if (!EMPLOYEE_EMAIL) {
//   console.warn(
//     "[mailer] EMPLOYEE_EMAIL (or SUJIT_EMAIL) is not set in .env — " +
//     "the employee will NOT receive renewal-reminder emails."
//   );
// }

// const transporter = nodemailer.createTransport({
//   host: process.env.SMTP_HOST,
//   port: Number(process.env.SMTP_PORT) || 587,
//   secure: process.env.SMTP_SECURE === "true",
//   auth: {
//     user: process.env.SMTP_USER,
//     pass: process.env.SMTP_PASS
//   },
//   tls: {
//     rejectUnauthorized: false,
//   },
// });

// transporter.verify((err) => {
//   if (err) {
//     console.error("[mailer] SMTP connection FAILED — reminder emails will not send:", err.message);
//   } else {
//     console.log("[mailer] SMTP connection verified OK");
//   }
// });

// function formatDate(date, pattern = DEFAULT_DATE_FORMAT) {
//   return date ? formatDateByPattern(date, pattern) : "N/A";
// }

// function buildReminderEmail(client, daysRemaining) {
//   const dayWord = daysRemaining === 1 ? "day" : "days";
//   const subject = `Reminder: License for ${client.partyName} expires in ${daysRemaining} ${dayWord}`;
//   const text =
//     `Hi ${client.partyName},\n\n` +
//     `This is a reminder that the license/registration for ${client.firmName || client.partyName} ` +
//     `is set to expire on ${formatDate(client.expiredDate, client.expiredDateFormat)} ` +
//     `(${daysRemaining} ${dayWord} from today).\n\n` +
//     `Please renew soon to avoid any interruption.\n\n` +
//     `Thank you.`;
//   const html = `
//     <div style="font-family: Arial, sans-serif; font-size: 15px; color: #222;">
//       <p>Hi ${client.partyName},</p>
//       <p>
//         This is a reminder that the license/registration for <strong>${client.firmName || client.partyName}</strong>
//         is set to expire on <strong>${formatDate(client.expiredDate, client.expiredDateFormat)}</strong>
//         (<strong>${daysRemaining} ${dayWord}</strong> from today).
//       </p>
//       <p>Please renew soon to avoid any interruption in service.</p>
//       <p>Thank you.</p>
//     </div>`;
//   return { subject, text, html };
// }

// function buildEmployeeReminderEmail(client, daysRemaining, employeeName) {
//   const dayWord = daysRemaining === 1 ? "day" : "days";
//   const subject = `Renewal Alert: ${client.partyName} expires in ${daysRemaining} ${dayWord}`;
//   const text =
//     `Hii ${employeeName},\n\n` +
//     `The party name is ${client.partyName}, their plan is expiring in just ${daysRemaining} ${dayWord}.\n\n` +
//     `Firm Name: ${client.firmName || "-"}\n` +
//     `Mobile No.: ${client.mobileNo || "-"}\n` +
//     `Client Email: ${client.emailId || "-"}\n` +
//     `License Number: ${client.licenseNumber || "-"}\n` +
//     `Expired Date: ${formatDate(client.expiredDate, client.expiredDateFormat)}\n\n` +
//     `Please follow up on the renewal.`;
//   const html = `
//     <div style="font-family: Arial, sans-serif; font-size: 15px; color: #222;">
//       <p>Hii ${employeeName},</p>
//       <p>
//         The party name is <strong>${client.partyName}</strong>, their plan is expiring in
//         just <strong>${daysRemaining} ${dayWord}</strong>.
//       </p>
//       <table style="border-collapse: collapse; margin-top: 10px;">
//         <tr><td style="padding:4px 10px 4px 0; color:#667085;">Firm Name</td><td>${client.firmName || "-"}</td></tr>
//         <tr><td style="padding:4px 10px 4px 0; color:#667085;">Mobile No.</td><td>${client.mobileNo || "-"}</td></tr>
//         <tr><td style="padding:4px 10px 4px 0; color:#667085;">Client Email</td><td>${client.emailId || "-"}</td></tr>
//         <tr><td style="padding:4px 10px 4px 0; color:#667085;">License Number</td><td>${client.licenseNumber || "-"}</td></tr>
//         <tr><td style="padding:4px 10px 4px 0; color:#667085;">Expired Date</td><td>${formatDate(client.expiredDate, client.expiredDateFormat)}</td></tr>
//       </table>
//       <p style="margin-top: 14px;">Please follow up on the renewal.</p>
//     </div>`;
//   return { subject, text, html };
// }

// async function sendReminderEmailToClient(client, daysRemaining) {
//   const { subject, text, html } = buildReminderEmail(client, daysRemaining);
//   await transporter.sendMail({
//     from: process.env.MAIL_FROM || process.env.SMTP_USER,
//     to: client.emailId,
//     subject,
//     text,
//     html
//   });
// }

// async function sendReminderEmailToEmployee(client, daysRemaining) {
//   const { subject, text, html } = buildEmployeeReminderEmail(client, daysRemaining, EMPLOYEE_NAME);
//   await transporter.sendMail({
//     from: process.env.MAIL_FROM || process.env.SMTP_USER,
//     to: EMPLOYEE_EMAIL,
//     subject,
//     text,
//     html
//   });
// }

// // ==================================================================
// // Core reminder job
// // ==================================================================
// // For each client:
// //   1. remaining = (expiryDate − today) in calendar days.
// //   2. milestone = currentMilestone(remaining) — the window they're in.
// //   3. If milestone is null, or already in remindersSent → skip.
// //   4. Otherwise: send one client email and one employee email,
// //      independently of each other. Then append the milestone to
// //      remindersSent so it doesn't re-fire today.
// async function runReminderEmailJob() {
//   const results = {
//     sent: [], skipped: [], failed: [],
//     employeeNotified: [], employeeFailed: []
//   };

//   const clients = await Client.find({ expiredDate: { $exists: true } });

//   for (const client of clients) {
//     const remaining = daysUntil(client.expiredDate);
//     const alreadySent = client.remindersSent || [];

//     const milestone = currentMilestone(remaining);
//     if (milestone === null) continue;
//     if (alreadySent.includes(milestone)) continue;

//     // 1) Client email (independent)
//     if (!client.emailId) {
//       results.skipped.push({
//         partyName: client.partyName,
//         milestone,
//         reason: "no email on file"
//       });
//     } else {
//       try {
//         await sendReminderEmailToClient(client, milestone);
//         results.sent.push({
//           partyName: client.partyName,
//           emailId: client.emailId,
//           daysRemaining: milestone
//         });
//       } catch (err) {
//         results.failed.push({
//           partyName: client.partyName,
//           emailId: client.emailId,
//           milestone,
//           error: err.message
//         });
//       }
//     }

//     // 2) Employee email — one per client, independent of client email
//     if (EMPLOYEE_EMAIL) {
//       try {
//         await sendReminderEmailToEmployee(client, milestone);
//         results.employeeNotified.push({
//           partyName: client.partyName,
//           employeeEmail: EMPLOYEE_EMAIL,
//           daysRemaining: milestone
//         });
//       } catch (err) {
//         results.employeeFailed.push({
//           partyName: client.partyName,
//           employeeEmail: EMPLOYEE_EMAIL,
//           milestone,
//           error: err.message
//         });
//       }
//     }

//     client.remindersSent = [...alreadySent, milestone];
//     await client.save();
//   }

//   return results;
// }

// // Manual trigger. Add `?force=1` to clear remindersSent on every client
// // first — useful after fixing the day calculation, or to re-fire the
// // current milestone for everyone.
// app.post("/api/reminders/send-emails", requireAuth, async (req, res) => {
//   try {
//     if (req.query.force === "1" || req.query.force === "true") {
//       await Client.updateMany({}, { $set: { remindersSent: [] } });
//       console.log("[reminders] force=1 — cleared remindersSent on all clients");
//     }
//     const results = await runReminderEmailJob();
//     res.json(results);
//   } catch (e) {
//     res.status(500).json({ message: e.message });
//   }
// });

// // Debug: shows exactly what the job sees per client, without sending.
// app.get("/api/reminders/debug", requireAuth, async (_, res) => {
//   try {
//     const clients = await Client.find({ expiredDate: { $exists: true } }).lean();
//     const now = new Date();
//     const rows = clients.map(c => {
//       const remaining = daysUntil(c.expiredDate);
//       const alreadySent = c.remindersSent || [];
//       const milestone = currentMilestone(remaining);
//       const willSend = milestone !== null && !alreadySent.includes(milestone);
//       return {
//         partyName: c.partyName,
//         emailId: c.emailId,
//         expiredDateISO: c.expiredDate,
//         expiredDateLocal: new Date(c.expiredDate).toString(),
//         todayLocal: now.toString(),
//         daysRemaining: remaining,
//         currentMilestone: milestone,
//         remindersSent: alreadySent,
//         wouldSendMilestone: willSend ? milestone : null,
//         wouldEmailClient: Boolean(willSend && c.emailId),
//         wouldEmailEmployee: Boolean(willSend && EMPLOYEE_EMAIL)
//       };
//     });
//     res.json({ employeeEmail: EMPLOYEE_EMAIL || null, reminderDays: EMAIL_REMINDER_DAYS, rows });
//   } catch (e) {
//     res.status(500).json({ message: e.message });
//   }
// });

// const reminderCronSchedule = process.env.REMINDER_CRON || "0 9 * * *";
// cron.schedule(reminderCronSchedule, () => {
//   console.log(`[reminders] Running scheduled reminder email job (${new Date().toISOString()})`);
//   runReminderEmailJob()
//     .then(results => {
//       console.log(
//         `[reminders] Done. Client emails — Sent: ${results.sent.length}, Skipped: ${results.skipped.length}, Failed: ${results.failed.length}. ` +
//         `Employee emails — Sent: ${results.employeeNotified.length}, Failed: ${results.employeeFailed.length}.`
//       );
//       results.sent.forEach(s =>
//         console.log(`[reminders]   CLIENT   -> ${s.partyName} <${s.emailId}> (${s.daysRemaining}d)`)
//       );
//       results.failed.forEach(f =>
//         console.error(`[reminders]   CLIENT FAILED -> ${f.partyName} <${f.emailId}> (${f.milestone}d): ${f.error}`)
//       );
//       results.skipped.forEach(s =>
//         console.warn(`[reminders]   SKIPPED -> ${s.partyName} (${s.milestone}d): ${s.reason}`)
//       );
//       results.employeeNotified.forEach(e =>
//         console.log(`[reminders]   EMPLOYEE -> ${EMPLOYEE_NAME} <${e.employeeEmail}> about ${e.partyName} (${e.daysRemaining}d)`)
//       );
//       results.employeeFailed.forEach(f =>
//         console.error(`[reminders]   EMPLOYEE FAILED -> <${f.employeeEmail}> about ${f.partyName} (${f.milestone}d): ${f.error}`)
//       );
//     })
//     .catch(err => console.error("[reminders] Job failed:", err.message));
// });

// // mail service end here

// // ==================================================================
// // Startup
// // ==================================================================
// mongoose.connect(process.env.MONGO_URI)
//   .then(async () => {
//     await ensureDefaultAdmin();
//     app.listen(port, () => console.log(`Backend running on http://localhost:${port}`));
//   })
//   .catch(err => {
//     console.error("MongoDB connection failed:", err.message);
//     process.exit(1);
//   });

// wrong - today is - 16th sep 26, expired date 21th sep 26 ( 17, 18, 19, 20, 21 = 4 days )
// right - today is - 16th sep 26, expired date 21th sep 26 ( 17, 18, 19, 20, 21 = 5 days )

// import express from "express";
// import cors from "cors";
// import dotenv from "dotenv";
// import mongoose from "mongoose";
// import XLSX from "xlsx";
// import nodemailer from "nodemailer";
// import cron from "node-cron";
// import bcrypt from "bcryptjs";
// import jwt from "jsonwebtoken";

// dotenv.config();

// const app = express();
// const port = process.env.PORT || 5000;

// app.use(cors({ origin: process.env.CORS_ORIGIN?.split(",") || "*" }));
// app.use(express.json());

// // ==================================================================
// // Auth config
// // ==================================================================
// const JWT_SECRET = process.env.JWT_SECRET || "please-change-this-secret-in-env";
// const TOKEN_EXPIRES_IN = process.env.TOKEN_EXPIRES_IN || "12h";
// const DEFAULT_ADMIN_USERNAME = process.env.ADMIN_USERNAME || "admin";
// const DEFAULT_ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "Admin@123";

// if (!process.env.JWT_SECRET) {
//   console.warn(
//     "[auth] WARNING: JWT_SECRET is not set in .env — using an insecure default. " +
//     "Set JWT_SECRET in backend/.env before using this in production."
//   );
// }

// // ==================================================================
// // Schemas
// // ==================================================================
// const userSchema = new mongoose.Schema({
//   username: { type: String, required: true, unique: true, trim: true, lowercase: true },
//   passwordHash: { type: String, required: true },
//   createdAt: { type: Date, default: Date.now }
// });
// const User = mongoose.model("User", userSchema);

// // Client record shape mirrors the manual-entry columns used by the dashboard:
// // PARTY NAME, FIRM NAME, USER ID, Password, Mobile No., Email.Id, License Number,
// // License Type, Client Number, Designation, Expired Date, K.O.B
// const clientSchema = new mongoose.Schema({
//   partyName: { type: String, required: true, trim: true },
//   firmName: { type: String, trim: true, default: "" },
//   userId: { type: String, trim: true, default: "" },
//   password: { type: String, default: "" },
//   mobileNo: { type: String, trim: true, default: "" },
//   emailId: { type: String, trim: true, default: "" },
//   licenseNumber: { type: String, trim: true, default: "" },
//   licenseType: { type: String, trim: true, default: "" },
//   clientNumber: { type: String, trim: true, default: "" },
//   designation: { type: String, trim: true, default: "" },
//   kob: { type: String, trim: true, default: "" },

//   expiredDate: { type: Date, required: true },
//   expiredDateFormat: { type: String, default: "dd-mm-yyyy" },

//   // Automatically set server-side the moment the "Renew" button is clicked.
//   lastRenewedAt: { type: Date, default: null },

//   renewed: { type: Boolean, default: false },
//   dismissedOn: { type: String, default: null },
//   // Tracks which day-milestone emails (15/10/5) have already been sent
//   // for the CURRENT expiry cycle, so a client is never emailed twice
//   // for the same milestone. Reset to [] whenever expiredDate changes.
//   remindersSent: { type: [Number], default: [] },

//   createdAt: { type: Date, default: Date.now },
//   updatedAt: { type: Date, default: Date.now }
// });

// const Client = mongoose.model("Client", clientSchema);

// const DEFAULT_DATE_FORMAT = "dd-mm-yyyy";

// function isValidDateParts(year, month, day) {
//   if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return false;
//   const d = new Date(year, month - 1, day);
//   return d.getFullYear() === year && d.getMonth() === month - 1 && d.getDate() === day;
// }

// // Accepts: Date objects, ISO strings from <input type="date">, and common
// // human text dates like dd-mm-yyyy / dd/mm/yyyy pasted in manually.
// function parseDate(value) {
//   if (value === null || value === undefined || value === "") return null;

//   if (value instanceof Date) {
//     return Number.isNaN(value.getTime())
//       ? null
//       : new Date(value.getFullYear(), value.getMonth(), value.getDate());
//   }

//   if (typeof value === "number") {
//     const parsed = XLSX.SSF.parse_date_code(value);
//     if (!parsed || !isValidDateParts(parsed.y, parsed.m, parsed.d)) return null;
//     return new Date(parsed.y, parsed.m - 1, parsed.d);
//   }

//   const text = String(value).trim();

//   let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
//   if (match) {
//     const [, y, m, d] = match.map(Number);
//     return isValidDateParts(y, m, d) ? new Date(y, m - 1, d) : null;
//   }

//   match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
//   if (match) {
//     let [, d, m, y] = match;
//     if (y.length === 2) y = `20${y}`;
//     d = Number(d); m = Number(m); y = Number(y);
//     return isValidDateParts(y, m, d) ? new Date(y, m - 1, d) : null;
//   }

//   const date = new Date(text);
//   return Number.isNaN(date.getTime())
//     ? null
//     : new Date(date.getFullYear(), date.getMonth(), date.getDate());
// }

// function normalizeDateFormat(format) {
//   if (!format || typeof format !== "string") return DEFAULT_DATE_FORMAT;
//   const cleaned = format.trim().replace(/\\\-/g, "-").replace(/\\\//g, "/");
//   const lower = cleaned.toLowerCase();
//   if (lower.includes("yy") && (lower.includes("dd") || lower.includes("d")) &&
//       (lower.includes("mm") || lower.includes("m"))) {
//     return cleaned;
//   }
//   return DEFAULT_DATE_FORMAT;
// }

// function formatDateByPattern(value, pattern = DEFAULT_DATE_FORMAT) {
//   const date = parseDate(value);
//   if (!date) return "";
//   const d = String(date.getDate()).padStart(2, "0");
//   const m = String(date.getMonth() + 1).padStart(2, "0");
//   const y = String(date.getFullYear());
//   const yy = y.slice(-2);

//   let out = pattern;
//   out = out.replace(/yyyy/gi, y).replace(/yy/gi, yy);
//   out = out.replace(/dd/gi, d);
//   out = out.replace(/mm/gi, m);
//   out = out.replace(/d/g, String(date.getDate()));
//   out = out.replace(/m/g, String(date.getMonth() + 1));
//   return out;
// }

// function dateKey(date = new Date()) {
//   return date.toISOString().slice(0, 10);
// }

// function daysUntil(date) {
//   const today = new Date();
//   today.setHours(0, 0, 0, 0);
//   const target = new Date(date);
//   target.setHours(0, 0, 0, 0);
//   return Math.round((target - today) / 86400000);
// }

// // Browser/extension reminder feed: anything expiring within this many days,
// // OR already overdue, is considered an "active" reminder. This intentionally
// // matches the dashboard's own "Expiring ≤15d" / "Expired" status badges —
// // previously this only matched exact 15/10/5/2-day milestones, so a client
// // with e.g. 6 days left showed a warning badge on the dashboard but never
// // appeared in the extension popup at all.
// const REMINDER_WINDOW_DAYS = 15;
// function reminderDaysLeft(client) {
//   const remaining = daysUntil(client.expiredDate);
//   return remaining <= REMINDER_WINDOW_DAYS ? remaining : null;
// }

// // ==================================================================
// // Auth helpers & middleware
// // ==================================================================
// function signToken(user) {
//   return jwt.sign({ sub: user._id.toString(), username: user.username }, JWT_SECRET, {
//     expiresIn: TOKEN_EXPIRES_IN
//   });
// }

// function requireAuth(req, res, next) {
//   const header = req.headers.authorization || "";
//   const [scheme, token] = header.split(" ");
//   if (scheme !== "Bearer" || !token) {
//     return res.status(401).json({ message: "Authentication required. Please log in." });
//   }
//   try {
//     const payload = jwt.verify(token, JWT_SECRET);
//     req.user = payload;
//     next();
//   } catch (e) {
//     return res.status(401).json({ message: "Session expired or invalid. Please log in again." });
//   }
// }

// async function ensureDefaultAdmin() {
//   const count = await User.countDocuments();
//   if (count > 0) return;
//   const passwordHash = await bcrypt.hash(DEFAULT_ADMIN_PASSWORD, 10);
//   await User.create({ username: DEFAULT_ADMIN_USERNAME.toLowerCase(), passwordHash });
//   console.log("=================================================================");
//   console.log(" No users found — a default admin account was created:");
//   console.log(`   Username: ${DEFAULT_ADMIN_USERNAME}`);
//   console.log(`   Password: ${DEFAULT_ADMIN_PASSWORD}`);
//   console.log(" Please log in and change this password immediately.");
//   console.log("=================================================================");
// }

// // ==================================================================
// // Auth routes (public)
// // ==================================================================
// app.get("/api/health", (_, res) => res.json({ ok: true }));

// app.post("/api/auth/login", async (req, res) => {
//   try {
//     const username = String(req.body.username || "").trim().toLowerCase();
//     const password = String(req.body.password || "");
//     if (!username || !password) {
//       return res.status(400).json({ message: "Username and password are required." });
//     }

//     const user = await User.findOne({ username });
//     if (!user) return res.status(401).json({ message: "Invalid username or password." });

//     const ok = await bcrypt.compare(password, user.passwordHash);
//     if (!ok) return res.status(401).json({ message: "Invalid username or password." });

//     const token = signToken(user);
//     res.json({ token, username: user.username });
//   } catch (e) {
//     res.status(500).json({ message: e.message });
//   }
// });

// app.get("/api/auth/me", requireAuth, (req, res) => {
//   res.json({ username: req.user.username });
// });

// app.post("/api/auth/change-password", requireAuth, async (req, res) => {
//   try {
//     const { currentPassword, newPassword } = req.body;
//     if (!currentPassword || !newPassword || newPassword.length < 6) {
//       return res.status(400).json({
//         message: "Current password and a new password (min 6 characters) are required."
//       });
//     }
//     const user = await User.findById(req.user.sub);
//     if (!user) return res.status(404).json({ message: "User not found." });

//     const ok = await bcrypt.compare(currentPassword, user.passwordHash);
//     if (!ok) return res.status(401).json({ message: "Current password is incorrect." });

//     user.passwordHash = await bcrypt.hash(newPassword, 10);
//     await user.save();
//     res.json({ message: "Password updated successfully." });
//   } catch (e) {
//     res.status(500).json({ message: e.message });
//   }
// });

// // ==================================================================
// // Client management routes (protected — manual entry only, no Excel import)
// // ==================================================================
// app.get("/api/clients", requireAuth, async (req, res) => {
//   try {
//     const search = String(req.query.search || "").trim();
//     const filter = search
//       ? { $or: [
//           { partyName: new RegExp(search, "i") },
//           { firmName: new RegExp(search, "i") },
//           { mobileNo: new RegExp(search, "i") },
//           { emailId: new RegExp(search, "i") },
//           { licenseNumber: new RegExp(search, "i") },
//           { clientNumber: new RegExp(search, "i") }
//         ] }
//       : {};
//     const clients = await Client.find(filter).sort({ expiredDate: 1 });
//     res.json(clients);
//   } catch (e) {
//     res.status(500).json({ message: e.message });
//   }
// });

// app.post("/api/clients", requireAuth, async (req, res) => {
//   try {
//     const expiredDate = parseDate(req.body.expiredDate);
//     if (!req.body.partyName || !String(req.body.partyName).trim()) {
//       return res.status(400).json({ message: "Party name is required." });
//     }
//     if (!expiredDate) {
//       return res.status(400).json({ message: "A valid expired date is required." });
//     }

//     const client = await Client.create({
//       partyName: req.body.partyName,
//       firmName: req.body.firmName,
//       userId: req.body.userId,
//       password: req.body.password,
//       mobileNo: req.body.mobileNo,
//       emailId: req.body.emailId,
//       licenseNumber: req.body.licenseNumber,
//       licenseType: req.body.licenseType,
//       clientNumber: req.body.clientNumber,
//       designation: req.body.designation,
//       kob: req.body.kob,
//       expiredDate,
//       expiredDateFormat: normalizeDateFormat(req.body.expiredDateFormat)
//     });
//     res.status(201).json(client);
//   } catch (e) {
//     res.status(400).json({ message: e.message });
//   }
// });

// app.put("/api/clients/:id", requireAuth, async (req, res) => {
//   try {
//     const existing = await Client.findById(req.params.id);
//     if (!existing) return res.status(404).json({ message: "Client not found" });

//     const update = {
//       partyName: req.body.partyName,
//       firmName: req.body.firmName,
//       userId: req.body.userId,
//       password: req.body.password,
//       mobileNo: req.body.mobileNo,
//       emailId: req.body.emailId,
//       licenseNumber: req.body.licenseNumber,
//       licenseType: req.body.licenseType,
//       clientNumber: req.body.clientNumber,
//       designation: req.body.designation,
//       kob: req.body.kob,
//       updatedAt: new Date()
//     };

//     if (Object.prototype.hasOwnProperty.call(req.body, "expiredDate")) {
//       const expiredDate = parseDate(req.body.expiredDate);
//       if (!expiredDate) return res.status(400).json({ message: "Valid expired date is required" });
//       update.expiredDate = expiredDate;
//       update.expiredDateFormat = normalizeDateFormat(
//         req.body.expiredDateFormat || existing.expiredDateFormat
//       );
//       // A changed expiry date starts a new reminder cycle.
//       if (existing.expiredDate?.getTime() !== expiredDate.getTime()) {
//         update.remindersSent = [];
//       }
//     }

//     if (!update.partyName || !String(update.partyName).trim()) {
//       return res.status(400).json({ message: "Party name is required." });
//     }

//     const client = await Client.findByIdAndUpdate(req.params.id, update, { new: true, runValidators: true });
//     res.json(client);
//   } catch (e) {
//     res.status(400).json({ message: e.message });
//   }
// });

// app.delete("/api/clients/:id", requireAuth, async (req, res) => {
//   try {
//     const client = await Client.findByIdAndDelete(req.params.id);
//     if (!client) return res.status(404).json({ message: "Client not found" });
//     res.json({ message: "Client deleted." });
//   } catch (e) {
//     res.status(400).json({ message: e.message });
//   }
// });

// // Renew: marks the client as renewed, optionally moves the expired date
// // forward, and — critically — ALWAYS stamps lastRenewedAt with the current
// // server date/time automatically. The client never types this in.
// app.post("/api/clients/:id/renew", requireAuth, async (req, res) => {
//   try {
//     const existing = await Client.findById(req.params.id);
//     if (!existing) return res.status(404).json({ message: "Client not found" });

//     const update = {
//       renewed: true,
//       lastRenewedAt: new Date(), // auto-set on click, never user-editable
//       dismissedOn: null,
//       updatedAt: new Date()
//     };

//     if (req.body.newExpiredDate) {
//       const newExpiredDate = parseDate(req.body.newExpiredDate);
//       if (!newExpiredDate) return res.status(400).json({ message: "Invalid new expired date." });
//       update.expiredDate = newExpiredDate;
//       update.expiredDateFormat = normalizeDateFormat(existing.expiredDateFormat);
//       if (existing.expiredDate?.getTime() !== newExpiredDate.getTime()) {
//         update.remindersSent = [];
//       }
//     }

//     const client = await Client.findByIdAndUpdate(req.params.id, update, { new: true });
//     res.json(client);
//   } catch (e) {
//     res.status(400).json({ message: e.message });
//   }
// });

// // ==================================================================
// // Reminder feed used by the Chrome extension popup.
// // Left un-authenticated on purpose: the browser extension has no login UI,
// // and this endpoint only exposes upcoming-expiry notices, not credentials.
// // ==================================================================
// app.post("/api/clients/:id/dismiss-today", async (req, res) => {
//   try {
//     const client = await Client.findByIdAndUpdate(
//       req.params.id,
//       { dismissedOn: dateKey(), updatedAt: new Date() },
//       { new: true }
//     );
//     res.json(client);
//   } catch (e) {
//     res.status(400).json({ message: e.message });
//   }
// });

// app.get("/api/reminders", async (_, res) => {
//   try {
//     const clients = await Client.find({ expiredDate: { $exists: true } });
//     const reminders = clients
//       .map(client => {
//         const remaining = reminderDaysLeft(client);
//         return {
//           ...client.toObject(),
//           daysRemaining: remaining,
//           dismissedToday: client.dismissedOn === dateKey()
//         };
//       })
//       .filter(client => client.daysRemaining !== null && !client.dismissedToday);
//     res.json(reminders);
//   } catch (e) {
//     res.status(500).json({ message: e.message });
//   }
// });

// // ==================================================================
// // Export to Excel (protected). Manual import has been removed — data only
// // ever enters the system through the dashboard's Add/Edit Client form.
// // ==================================================================
// app.get("/api/export/excel", requireAuth, async (_, res) => {
//   try {
//     const clients = await Client.find({}).sort({ expiredDate: 1 }).lean();
//     const headers = [
//       "party_name", "firm_name", "user_id", "password", "mobile_no", "email_id",
//       "license_number", "license_type", "client_number", "designation", "kob",
//       "expired_date", "last_renewed_at"
//     ];
//     const data = clients.map(c => ({
//       party_name: c.partyName ?? "",
//       firm_name: c.firmName ?? "",
//       user_id: c.userId ?? "",
//       password: c.password ?? "",
//       mobile_no: c.mobileNo ?? "",
//       email_id: c.emailId ?? "",
//       license_number: c.licenseNumber ?? "",
//       license_type: c.licenseType ?? "",
//       client_number: c.clientNumber ?? "",
//       designation: c.designation ?? "",
//       kob: c.kob ?? "",
//       expired_date: c.expiredDate ? new Date(c.expiredDate) : "",
//       last_renewed_at: c.lastRenewedAt ? new Date(c.lastRenewedAt) : ""
//     }));

//     // IMPORTANT: cellDates:true makes json_to_sheet store real JS Date values
//     // as proper "d"-typed cells. Without this, dates are converted to plain
//     // Excel serial numbers ("n"-typed); forcing those cells to type "d"
//     // afterwards (without cellDates:true) makes the writer re-interpret the
//     // serial number as if it were a raw JS millisecond timestamp, producing
//     // bogus 01-01-1970 dates in the downloaded file. cellDates:true avoids
//     // that entirely.
//     const ws = XLSX.utils.json_to_sheet(data, { header: headers, cellDates: true });
//     clients.forEach((c, i) => {
//       const row = i + 2;
//       if (c.expiredDate && ws[`L${row}`]) {
//         ws[`L${row}`].z = normalizeDateFormat(c.expiredDateFormat);
//       }
//       if (c.lastRenewedAt && ws[`M${row}`]) {
//         ws[`M${row}`].z = "dd-mm-yyyy hh:mm";
//       }
//     });

//     const wb = XLSX.utils.book_new();
//     XLSX.utils.book_append_sheet(wb, ws, "Clients");
//     const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx", cellDates: true });
//     res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//     res.setHeader("Content-Disposition", 'attachment; filename="clients.xlsx"');
//     res.send(buffer);
//   } catch (e) {
//     res.status(500).json({ message: e.message });
//   }
// });

// // mail service start here

// // ==================================================================
// // Email reminders (Nodemailer)
// // ==================================================================
// const EMAIL_REMINDER_DAYS = [15, 10, 5];

// // A fixed employee inbox that gets notified alongside the client on every
// // reminder milestone — this is the person who actually handles renewals
// // for these clients, not the client themselves.
// // EMPLOYEE_EMAIL is the preferred variable name; SUJIT_EMAIL is kept as a
// // fallback since it was already present in this project's .env.
// const EMPLOYEE_NAME = process.env.EMPLOYEE_NAME || "Team";
// const EMPLOYEE_EMAIL = process.env.EMPLOYEE_EMAIL || process.env.SUJIT_EMAIL || "";

// if (!EMPLOYEE_EMAIL) {
//   console.warn(
//     "[mailer] EMPLOYEE_EMAIL (or SUJIT_EMAIL) is not set in .env — " +
//     "the employee will NOT receive renewal-reminder emails."
//   );
// }

// const transporter = nodemailer.createTransport({
//   host: process.env.SMTP_HOST,
//   port: Number(process.env.SMTP_PORT) || 587,
//   secure: process.env.SMTP_SECURE === "true",
//   auth: {
//     user: process.env.SMTP_USER,
//     pass: process.env.SMTP_PASS
//   },
//   tls: {
//     rejectUnauthorized: false,
//   },
// });

// transporter.verify((err) => {
//   if (err) {
//     console.error("[mailer] SMTP connection FAILED — reminder emails will not send:", err.message);
//   } else {
//     console.log("[mailer] SMTP connection verified OK");
//   }
// });

// function formatDate(date, pattern = DEFAULT_DATE_FORMAT) {
//   return date ? formatDateByPattern(date, pattern) : "N/A";
// }

// function buildReminderEmail(client, daysRemaining) {
//   const subject = `Reminder: License for ${client.partyName} expires in ${daysRemaining} day${daysRemaining === 1 ? "" : "s"}`;
//   const text =
//     `Hi ${client.partyName},\n\n` +
//     `This is a reminder that the license/registration for ${client.firmName || client.partyName} ` +
//     `is set to expire on ${formatDate(client.expiredDate, client.expiredDateFormat)} ` +
//     `(${daysRemaining} day${daysRemaining === 1 ? "" : "s"} from today).\n\n` +
//     `Please renew soon to avoid any interruption.\n\n` +
//     `Thank you.`;
//   const html = `
//     <div style="font-family: Arial, sans-serif; font-size: 15px; color: #222;">
//       <p>Hi ${client.partyName},</p>
//       <p>
//         This is a reminder that the license/registration for <strong>${client.firmName || client.partyName}</strong>
//         is set to expire on <strong>${formatDate(client.expiredDate, client.expiredDateFormat)}</strong>
//         (<strong>${daysRemaining} day${daysRemaining === 1 ? "" : "s"}</strong> from today).
//       </p>
//       <p>Please renew soon to avoid any interruption in service.</p>
//       <p>Thank you.</p>
//     </div>`;
//   return { subject, text, html };
// }

// // The fixed employee notification. Wording follows exactly what was asked
// // for: "Hii <employee name>, the party name is <party>, their plan is
// // expiring in just <days> days" — plus a few extra details so the employee
// // can act on it without opening the dashboard first.
// function buildEmployeeReminderEmail(client, daysRemaining, employeeName) {
//   const dayWord = daysRemaining === 1 ? "day" : "days";
//   const subject = `Renewal Alert: ${client.partyName} expires in ${daysRemaining} ${dayWord}`;
//   const text =
//     `Hii ${employeeName},\n\n` +
//     `The party name is ${client.partyName}, their plan is expiring in just ${daysRemaining} ${dayWord}.\n\n` +
//     `Firm Name: ${client.firmName || "-"}\n` +
//     `Mobile No.: ${client.mobileNo || "-"}\n` +
//     `Client Email: ${client.emailId || "-"}\n` +
//     `License Number: ${client.licenseNumber || "-"}\n` +
//     `Expired Date: ${formatDate(client.expiredDate, client.expiredDateFormat)}\n\n` +
//     `Please follow up on the renewal.`;
//   const html = `
//     <div style="font-family: Arial, sans-serif; font-size: 15px; color: #222;">
//       <p>Hii ${employeeName},</p>
//       <p>
//         The party name is <strong>${client.partyName}</strong>, their plan is expiring in
//         just <strong>${daysRemaining} ${dayWord}</strong>.
//       </p>
//       <table style="border-collapse: collapse; margin-top: 10px;">
//         <tr><td style="padding:4px 10px 4px 0; color:#667085;">Firm Name</td><td>${client.firmName || "-"}</td></tr>
//         <tr><td style="padding:4px 10px 4px 0; color:#667085;">Mobile No.</td><td>${client.mobileNo || "-"}</td></tr>
//         <tr><td style="padding:4px 10px 4px 0; color:#667085;">Client Email</td><td>${client.emailId || "-"}</td></tr>
//         <tr><td style="padding:4px 10px 4px 0; color:#667085;">License Number</td><td>${client.licenseNumber || "-"}</td></tr>
//         <tr><td style="padding:4px 10px 4px 0; color:#667085;">Expired Date</td><td>${formatDate(client.expiredDate, client.expiredDateFormat)}</td></tr>
//       </table>
//       <p style="margin-top: 14px;">Please follow up on the renewal.</p>
//     </div>`;
//   return { subject, text, html };
// }

// async function sendReminderEmailToClient(client, daysRemaining) {
//   const { subject, text, html } = buildReminderEmail(client, daysRemaining);
//   await transporter.sendMail({
//     from: process.env.MAIL_FROM || process.env.SMTP_USER,
//     to: client.emailId,
//     subject,
//     text,
//     html
//   });
// }

// async function sendReminderEmailToEmployee(client, daysRemaining) {
//   const { subject, text, html } = buildEmployeeReminderEmail(client, daysRemaining, EMPLOYEE_NAME);
//   await transporter.sendMail({
//     from: process.env.MAIL_FROM || process.env.SMTP_USER,
//     to: EMPLOYEE_EMAIL,
//     subject,
//     text,
//     html
//   });
// }

// async function runReminderEmailJob() {
//   const results = {
//     sent: [], skipped: [], failed: [],
//     employeeNotified: [], employeeFailed: []
//   };
//   // NOTE: intentionally does NOT filter by `renewed`. The dashboard and the
//   // Chrome-extension reminder popup both decide "is this client due a
//   // reminder?" purely from expiredDate (see statusFor() in Dashboard.jsx and
//   // reminderDaysLeft() above) — the `renewed` flag just marks that someone
//   // clicked the Renew button, it does not by itself mean the expiry date
//   // moved forward. Filtering on `renewed` here caused clients who were
//   // marked renewed without picking a new expiry date to be silently
//   // excluded from reminder emails while still showing as "active" in the
//   // dashboard/extension. remindersSent (deduped per milestone, reset
//   // whenever expiredDate actually changes) is what prevents duplicate
//   // emails — that's enough on its own.
//   const clients = await Client.find({ expiredDate: { $exists: true } });

//   for (const client of clients) {
//     const remaining = daysUntil(client.expiredDate);
//     if (!EMAIL_REMINDER_DAYS.includes(remaining)) continue;
//     if ((client.remindersSent || []).includes(remaining)) continue;

//     // 1) Email the client, if they have one on file.
//     if (!client.emailId) {
//       results.skipped.push({ partyName: client.partyName, reason: "no email on file" });
//     } else {
//       try {
//         await sendReminderEmailToClient(client, remaining);
//         results.sent.push({ partyName: client.partyName, emailId: client.emailId, daysRemaining: remaining });
//       } catch (err) {
//         results.failed.push({ partyName: client.partyName, emailId: client.emailId, error: err.message });
//       }
//     }

//     // 2) ALSO email the fixed employee inbox — independent of whether the
//     // client's own email exists or succeeded, so the employee always finds
//     // out a renewal is coming up.
//     if (EMPLOYEE_EMAIL) {
//       try {
//         await sendReminderEmailToEmployee(client, remaining);
//         results.employeeNotified.push({ partyName: client.partyName, employeeEmail: EMPLOYEE_EMAIL, daysRemaining: remaining });
//       } catch (err) {
//         results.employeeFailed.push({ partyName: client.partyName, employeeEmail: EMPLOYEE_EMAIL, error: err.message });
//       }
//     }

//     // Milestone handled — don't re-notify for this same day-count again.
//     client.remindersSent = [...(client.remindersSent || []), remaining];
//     await client.save();
//   }

//   return results;
// }

// app.post("/api/reminders/send-emails", requireAuth, async (_, res) => {
//   try {
//     const results = await runReminderEmailJob();
//     res.json(results);
//   } catch (e) {
//     res.status(500).json({ message: e.message });
//   }
// });

// const reminderCronSchedule = process.env.REMINDER_CRON || "0 9 * * *";
// cron.schedule(reminderCronSchedule, () => {
//   console.log(`[reminders] Running scheduled reminder email job (${new Date().toISOString()})`);
//   runReminderEmailJob()
//     .then(results => {
//       console.log(
//         `[reminders] Done. Client emails — Sent: ${results.sent.length}, Skipped: ${results.skipped.length}, Failed: ${results.failed.length}. ` +
//         `Employee emails — Sent: ${results.employeeNotified.length}, Failed: ${results.employeeFailed.length}.`
//       );
//       results.failed.forEach(f =>
//         console.error(`[reminders]   CLIENT EMAIL FAILED -> ${f.partyName} <${f.emailId}>: ${f.error}`)
//       );
//       results.skipped.forEach(s =>
//         console.warn(`[reminders]   SKIPPED -> ${s.partyName}: ${s.reason}`)
//       );
//       results.employeeFailed.forEach(f =>
//         console.error(`[reminders]   EMPLOYEE EMAIL FAILED -> ${f.partyName} -> <${f.employeeEmail}>: ${f.error}`)
//       );
//     })
//     .catch(err => console.error("[reminders] Job failed:", err.message));
// });

// // mail service end here

// // ==================================================================
// // Startup
// // ==================================================================
// mongoose.connect(process.env.MONGO_URI)
//   .then(async () => {
//     await ensureDefaultAdmin();
//     app.listen(port, () => console.log(`Backend running on http://localhost:${port}`));
//   })
//   .catch(err => {
//     console.error("MongoDB connection failed:", err.message);
//     process.exit(1);
//   });
