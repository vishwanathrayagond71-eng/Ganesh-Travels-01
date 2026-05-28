// ============================================================
// server.js - Main Express Application Entry Point
// Tours and Tour - Professional Tourism Website
// ============================================================

require('dotenv').config(); // Load environment variables from .env

const express = require('express');
const ejs = require('ejs'); // Force Netlify bundler to package EJS
const session = require('express-session');
const MongoStore = require('connect-mongo');
const flash = require('connect-flash');
const path = require('path');
const excel = require('./utils/excelService');
const { initAllSheets } = excel;

const app = express();
const PORT = process.env.PORT || 3000;

// ============================================================
// VIEW ENGINE SETUP - EJS Templates
// ============================================================
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// ============================================================
// STATIC FILES - CSS, JS, Images
// ============================================================
app.use(express.static(path.join(__dirname, 'public')));

// ============================================================
// BODY PARSING MIDDLEWARE
// ============================================================
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// ============================================================
// SESSION MIDDLEWARE
// ============================================================
const sessionOptions = {
  secret: process.env.SESSION_SECRET || 'tours_secret_key',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 24 * 60 * 60 * 1000 } // 24 hours
};

if (process.env.MONGODB_URI) {
  sessionOptions.store = MongoStore.create({
    mongoUrl: process.env.MONGODB_URI,
    collectionName: 'sessions',
    ttl: 24 * 60 * 60 // 1 day
  });
}

app.use(session(sessionOptions));

// ============================================================
// FLASH MESSAGES MIDDLEWARE
// ============================================================
app.use(flash());

// ============================================================
// GLOBAL VARIABLES MIDDLEWARE
// ============================================================
app.use(async (req, res, next) => {
  res.locals.user = req.session ? (req.session.user || null) : null;
  res.locals.admin = req.session ? (req.session.admin || null) : null;
  try {
    const phoneSetting = await excel.findOne('settings', 'key', 'contact_phone');
    const emailSetting = await excel.findOne('settings', 'key', 'contact_email');
    res.locals.contactPhone = phoneSetting ? phoneSetting.value : '+91 98765 43210';
    res.locals.contactEmail = emailSetting ? emailSetting.value : 'aditya777@gmail.com';
  } catch (err) {
    res.locals.contactPhone = '+91 98765 43210';
    res.locals.contactEmail = 'aditya777@gmail.com';
  }
  next();
});

// ============================================================
// ROUTES
// ============================================================
const indexRoutes = require('./routes/index');
app.use('/', indexRoutes);

// ============================================================
// 404 PAGE
// ============================================================
app.use((req, res) => {
  res.status(404).render('404', { title: '404 - Page Not Found', user: req.session.user || null });
});

// ============================================================
// ERROR HANDLER
// ============================================================
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).send('Something went wrong! Please try again.');
});

// ============================================================
// START SERVER
// ============================================================
async function startServer() {
  try {
    // Initialize all Excel sheets before starting the server
    await initAllSheets();
    app.listen(PORT, () => {
      console.log(`🚀 Ganesh Travels is running at http://localhost:${PORT}`);
      console.log(`📊 Admin Panel: http://localhost:${PORT}/admin/login`);
    });
  } catch (err) {
    console.error('Failed to start server:', err);
    process.exit(1);
  }
}

if (require.main === module) {
  startServer();
}

module.exports = app;
