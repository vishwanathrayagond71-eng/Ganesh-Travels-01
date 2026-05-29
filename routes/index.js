// ============================================================
// routes/index.js - All Application Routes
// Handles public pages, auth, user dashboard, admin panel
// ============================================================

const express = require('express');
const router = express.Router();
const bcrypt = require('bcryptjs');
const { v4: uuidv4 } = require('uuid');
const excel = require('../utils/excelService');
const { requireUser, requireAdmin, redirectIfLoggedIn, redirectIfAdminLoggedIn } = require('../middleware/auth');
const path = require('path');

// Middleware moved to server.js globally

// ============================================================
// DATA: Destinations, Packages (Static data - no DB needed)
// ============================================================
const destinations = [
  { id: 1, name: 'Taj Mahal, Agra', category: 'historical', country: 'India', image: 'https://images.unsplash.com/photo-1564507592333-c60657eea523?w=600', description: 'One of the Seven Wonders of the World — a magnificent marble mausoleum in Agra.', price: 8999, rating: 4.9, reviews: 1240 },
  { id: 2, name: 'Goa Beaches', category: 'beach', country: 'India', image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=600', description: 'Sun-kissed beaches, vibrant nightlife, and Portuguese heritage await you in Goa.', price: 12999, rating: 4.7, reviews: 890 },
  { id: 3, name: 'Manali, Himachal', category: 'hill', country: 'India', image: 'https://images.unsplash.com/photo-1626621341517-bbf3d9990a23?w=600', description: 'Scenic hill station surrounded by snow-capped Himalayas, perfect for adventure seekers.', price: 14999, rating: 4.8, reviews: 760 },
  { id: 4, name: 'Jaipur, Rajasthan', category: 'historical', country: 'India', image: 'https://images.unsplash.com/photo-1599661046289-e31897846e41?w=600', description: 'The Pink City with majestic forts, palaces and vibrant bazaars of royal Rajasthan.', price: 9999, rating: 4.6, reviews: 654 },
  { id: 5, name: 'Kerala Backwaters', category: 'nature', country: 'India', image: 'https://images.unsplash.com/photo-1602216056096-3b40cc0c9944?w=600', description: 'Serene houseboat rides through lush green backwaters and coconut-fringed shores.', price: 16999, rating: 4.9, reviews: 920 },
  { id: 6, name: 'Varanasi, UP', category: 'spiritual', country: 'India', image: 'https://images.unsplash.com/photo-1561361058-c24e01e34f44?w=600', description: 'The spiritual capital of India — ancient ghats, sacred rituals, and divine Ganga Aarti.', price: 7999, rating: 4.5, reviews: 540 },
  { id: 7, name: 'Ladakh, J&K', category: 'hill', country: 'India', image: 'https://images.unsplash.com/photo-1536867114-e8b98cca2473?w=600', description: 'High-altitude desert landscape with monasteries, azure lakes and rugged mountains.', price: 24999, rating: 4.9, reviews: 430 },
  { id: 8, name: 'Andaman Islands', category: 'beach', country: 'India', image: 'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=600', description: 'Crystal-clear waters, pristine white-sand beaches and vibrant coral reefs.', price: 29999, rating: 4.8, reviews: 380 },
  { id: 9, name: 'Paris, France', category: 'international', country: 'France', image: 'https://images.unsplash.com/photo-1502602898657-3e91760cbb34?w=600', description: 'The City of Love — Eiffel Tower, world-class art, cuisine and fashion await.', price: 89999, rating: 4.9, reviews: 2100 },
  { id: 10, name: 'Bali, Indonesia', category: 'international', country: 'Indonesia', image: 'https://images.unsplash.com/photo-1537996194471-e657df975ab4?w=600', description: 'Island of Gods with lush rice terraces, Hindu temples, surf beaches and luxury resorts.', price: 54999, rating: 4.8, reviews: 1560 },
  { id: 11, name: 'Dubai, UAE', category: 'international', country: 'UAE', image: 'https://images.unsplash.com/photo-1512453979798-5ea266f8880c?w=600', description: 'Ultra-modern skyline, luxury shopping, desert safaris and world-record attractions.', price: 79999, rating: 4.7, reviews: 980 },
  { id: 12, name: 'Maldives', category: 'beach', country: 'Maldives', image: 'https://images.unsplash.com/photo-1514282401047-d79a71a590e8?w=600', description: 'Paradise overwater bungalows, turquoise lagoons and exceptional marine life.', price: 129999, rating: 5.0, reviews: 670 }
];

const packages = [
  { id: 1, name: 'Golden Triangle Family Tour', type: 'family', image: 'https://images.unsplash.com/photo-1599661046289-e31897846e41?w=600', price: 34999, duration: '7 Days / 6 Nights', destinations: 'Delhi - Agra - Jaipur', services: ['Hotel (4★)', 'Breakfast & Dinner', 'AC Transport', 'Tour Guide', 'Entry Tickets'], rating: 4.8 },
  { id: 2, name: 'Kerala Honeymoon Special', type: 'honeymoon', image: 'https://images.unsplash.com/photo-1602216056096-3b40cc0c9944?w=600', price: 54999, duration: '6 Days / 5 Nights', destinations: 'Munnar - Thekkady - Alleppey', services: ['Luxury Hotel', 'All Meals', 'Houseboat Stay', 'Couple Spa', 'Romantic Dinner'], rating: 4.9 },
  { id: 3, name: 'Ladakh Adventure Expedition', type: 'adventure', image: 'https://images.unsplash.com/photo-1536867114-e8b98cca2473?w=600', price: 44999, duration: '9 Days / 8 Nights', destinations: 'Leh - Nubra - Pangong', services: ['Camp Stay', 'Breakfast', 'Bike Rental', 'Adventure Guide', 'Permits'], rating: 4.9 },
  { id: 4, name: 'Goa Solo Beach Getaway', type: 'solo', image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=600', price: 18999, duration: '4 Days / 3 Nights', destinations: 'North Goa - South Goa', services: ['Boutique Hotel', 'Breakfast', 'Scooter Rental', 'Beach Activities', 'City Tour'], rating: 4.6 },
  { id: 5, name: 'Rajasthan Group Heritage Tour', type: 'group', image: 'https://images.unsplash.com/photo-1477587458883-47145ed68d6a?w=600', price: 28999, duration: '8 Days / 7 Nights', destinations: 'Jaipur - Jodhpur - Udaipur', services: ['Heritage Hotel', 'Daily Breakfast', 'AC Bus', 'Camel Safari', 'Folk Show'], rating: 4.7 },
  { id: 6, name: 'Bali International Package', type: 'international', image: 'https://images.unsplash.com/photo-1537996194471-e657df975ab4?w=600', price: 69999, duration: '7 Days / 6 Nights', destinations: 'Kuta - Ubud - Nusa Dua', services: ['5★ Resort', 'All Meals', 'Flights', 'Spa & Wellness', 'Temple Tours'], rating: 4.8 }
];

// ============================================================
// PUBLIC PAGES
// ============================================================

// Home Page
router.get('/', async (req, res) => {
  try {
    const reviews = await excel.readData('reviews');
    const approvedReviews = reviews.filter(r => r.status === 'approved').slice(-6);
    res.render('home', {
      title: 'Ganesh Travels - Discover India & Beyond',
      user: req.session.user || null,
      destinations: destinations.slice(0, 6),
      packages: packages.slice(0, 3),
      reviews: approvedReviews,
      messages: req.flash()
    });
  } catch (err) {
    console.error(err);
    res.render('home', { title: 'Ganesh Travels', user: req.session.user || null, destinations: destinations.slice(0, 6), packages: packages.slice(0, 3), reviews: [], messages: {} });
  }
});

// About Page
router.get('/about', async (req, res) => {
  try {
    const team = await excel.readData('team');
    res.render('about', { title: 'About Us - Ganesh Travels', user: req.session.user || null, team, messages: req.flash() });
  } catch (err) {
    res.render('about', { title: 'About Us - Ganesh Travels', user: req.session.user || null, team: [], messages: req.flash() });
  }
});

// Destinations Page
router.get('/destinations', (req, res) => {
  res.render('destinations', { title: 'Destinations - Ganesh Travels', user: req.session.user || null, destinations, messages: req.flash() });
});

// Packages Page
router.get('/packages', (req, res) => {
  res.render('packages', { title: 'Tour Packages - Ganesh Travels', user: req.session.user || null, packages, messages: req.flash() });
});

// Contact Page
router.get('/contact', (req, res) => {
  res.render('contact', { title: 'Contact Us - Ganesh Travels', user: req.session.user || null, messages: req.flash() });
});

// Save Contact Form
router.post('/contact', async (req, res) => {
  try {
    const { name, email, subject, message } = req.body;
    if (!name || !email || !message) {
      req.flash('error', 'Please fill all required fields.');
      return res.redirect('/contact');
    }
    await excel.addRow('contacts', { id: uuidv4(), name, email, subject: subject || 'General Inquiry', message, createdAt: new Date().toLocaleString() });
    req.flash('success', 'Your message has been sent! We will contact you soon.');
    res.redirect('/contact');
  } catch (err) {
    console.error(err);
    req.flash('error', 'Something went wrong. Please try again.');
    res.redirect('/contact');
  }
});

// Newsletter Subscribe
router.post('/newsletter', async (req, res) => {
  try {
    const { email } = req.body;
    if (!email) return res.json({ success: false, message: 'Email is required.' });
    const existing = await excel.findOne('newsletter', 'email', email);
    if (existing) return res.json({ success: false, message: 'You are already subscribed!' });
    await excel.addRow('newsletter', { id: uuidv4(), email, createdAt: new Date().toLocaleString() });
    res.json({ success: true, message: 'You have been subscribed successfully! 🎉' });
  } catch (err) {
    console.error(err);
    res.json({ success: false, message: 'Something went wrong.' });
  }
});

// Booking Page (GET)
router.get('/booking', (req, res) => {
  const preDestination = req.query.destination || '';
  res.render('booking', { title: 'Book Your Tour - Ganesh Travels', user: req.session.user || null, destinations, packages, preDestination, messages: req.flash() });
});

// Booking Submit (POST)
router.post('/booking', async (req, res) => {
  try {
    const { destination, date, guests, name, email, phone, message } = req.body;
    if (!destination || !date || !guests || !name || !email || !phone) {
      req.flash('error', 'Please fill all required booking fields.');
      return res.redirect('/booking');
    }
    const userId = req.session.user ? req.session.user.id : 'guest';
    const userEmail = req.session.user ? req.session.user.email : email;
    await excel.addRow('bookings', {
      id: uuidv4(), userId, userName: name, userEmail, destination, date, guests,
      name, email, phone, message: message || '', status: 'pending', createdAt: new Date().toLocaleString()
    });
    req.flash('success', `Booking confirmed! 🎉 We'll contact you at ${email} soon.`);
    res.redirect('/booking');
  } catch (err) {
    console.error(err);
    req.flash('error', 'Booking failed. Please try again.');
    res.redirect('/booking');
  }
});

// Reviews Page (GET)
router.get('/reviews', async (req, res) => {
  try {
    const reviews = await excel.readData('reviews');
    const approvedReviews = reviews.filter(r => r.status === 'approved');
    res.render('reviews', { title: 'Customer Reviews - Ganesh Travels', user: req.session.user || null, reviews: approvedReviews, destinations, messages: req.flash() });
  } catch (err) {
    res.render('reviews', { title: 'Reviews', user: req.session.user || null, reviews: [], destinations, messages: {} });
  }
});

// Submit Review (POST)
router.post('/reviews', async (req, res) => {
  try {
    const { destination, rating, comment } = req.body;
    if (!destination || !rating || !comment) {
      req.flash('error', 'Please fill all review fields.');
      return res.redirect('/reviews');
    }
    const userName = req.session.user ? req.session.user.name : req.body.name || 'Anonymous';
    const userEmail = req.session.user ? req.session.user.email : req.body.email || 'guest@guest.com';
    await excel.addRow('reviews', {
      id: uuidv4(), userEmail, userName, destination, rating, comment, status: 'approved', createdAt: new Date().toLocaleString()
    });
    req.flash('success', 'Thank you for your review! 🌟');
    res.redirect('/reviews');
  } catch (err) {
    console.error(err);
    req.flash('error', 'Could not submit review. Try again.');
    res.redirect('/reviews');
  }
});

// ============================================================
// USER AUTHENTICATION ROUTES
// ============================================================

// Register (GET)
router.get('/register', redirectIfLoggedIn, (req, res) => {
  res.render('register', { title: 'Register - Ganesh Travels', user: null, messages: req.flash() });
});

// Register (POST)
router.post('/register', redirectIfLoggedIn, async (req, res) => {
  try {
    const { name, email, phone, password, confirmPassword } = req.body;

    // Validate input
    if (!name || !email || !phone || !password || !confirmPassword) {
      req.flash('error', 'All fields are required.');
      return res.redirect('/register');
    }
    
    // Check if it is a Gmail ID
    if (!email.toLowerCase().endsWith('@gmail.com')) {
      req.flash('error', 'Only Gmail addresses (@gmail.com) are allowed for registration.');
      return res.redirect('/register');
    }

    if (password !== confirmPassword) {
      req.flash('error', 'Passwords do not match.');
      return res.redirect('/register');
    }
    if (password.length < 6) {
      req.flash('error', 'Password must be at least 6 characters.');
      return res.redirect('/register');
    }

    // Check if email already exists
    const existing = await excel.findOne('users', 'email', email);
    if (existing) {
      req.flash('error', 'This Gmail ID is already registered. Please login.');
      return res.redirect('/register');
    }

    // Hash password and save user
    const hashedPassword = await bcrypt.hash(password, 10);
    const userId = uuidv4();
    await excel.addRow('users', { id: userId, name, email, phone, password: hashedPassword, createdAt: new Date().toLocaleString() });

    req.flash('success', 'Registration successful! Please login.');
    res.redirect('/login');
  } catch (err) {
    console.error(err);
    req.flash('error', 'Registration failed. Please try again.');
    res.redirect('/register');
  }
});

// Login (GET)
router.get('/login', redirectIfLoggedIn, (req, res) => {
  res.render('login', { title: 'Login - Ganesh Travels', user: null, messages: req.flash() });
});

// Login (POST)
router.post('/login', redirectIfLoggedIn, async (req, res) => {
  try {
    const { email, password } = req.body;
    if (!email || !password) {
      req.flash('error', 'Email and password are required.');
      return res.redirect('/login');
    }

    const user = await excel.findOne('users', 'email', email);
    if (!user) {
      req.flash('error', 'Invalid email or password.');
      return res.redirect('/login');
    }

    const isMatch = await bcrypt.compare(password, user.password);
    if (!isMatch) {
      req.flash('error', 'Invalid email or password.');
      return res.redirect('/login');
    }

    // Save user in session
    req.session.user = { id: user.id, name: user.name, email: user.email, phone: user.phone };
    req.flash('success', `Welcome back, ${user.name}! 🎉`);
    res.redirect('/dashboard');
  } catch (err) {
    console.error(err);
    req.flash('error', 'Login failed. Please try again.');
    res.redirect('/login');
  }
});

// Logout
router.get('/logout', (req, res) => {
  req.session.destroy();
  res.redirect('/login');
});

// ============================================================
// USER DASHBOARD ROUTES
// ============================================================

router.get('/dashboard', requireUser, async (req, res) => {
  try {
    const allBookings = await excel.readData('bookings');
    const allReviews = await excel.readData('reviews');
    const userBookings = allBookings.filter(b => b.userEmail === req.session.user.email);
    const userReviews = allReviews.filter(r => r.userEmail === req.session.user.email);
    res.render('dashboard', {
      title: 'My Dashboard - Ganesh Travels',
      user: req.session.user,
      bookings: userBookings,
      reviews: userReviews,
      messages: req.flash()
    });
  } catch (err) {
    console.error(err);
    res.render('dashboard', { title: 'Dashboard', user: req.session.user, bookings: [], reviews: [], messages: {} });
  }
});

// Edit Profile
router.post('/dashboard/edit-profile', requireUser, async (req, res) => {
  try {
    const { name, phone } = req.body;
    await excel.updateRow('users', req.session.user.id, { name, phone });
    req.session.user.name = name;
    req.session.user.phone = phone;
    req.flash('success', 'Profile updated successfully!');
    res.redirect('/dashboard');
  } catch (err) {
    req.flash('error', 'Failed to update profile.');
    res.redirect('/dashboard');
  }
});

// ============================================================
// ADMIN ROUTES
// ============================================================

// Admin Login (GET)
router.get('/admin/login', redirectIfAdminLoggedIn, (req, res) => {
  res.render('admin-login', { title: 'Admin Login - Ganesh Travels', messages: req.flash() });
});

// Admin Login (POST)
router.post('/admin/login', redirectIfAdminLoggedIn, (req, res) => {
  const { email, password } = req.body;
  const adminEmail = process.env.ADMIN_EMAIL;
  const adminPassword = process.env.ADMIN_PASSWORD;

  if (!adminEmail || !adminPassword) {
    req.flash('error', 'Admin credentials are not configured on the server.');
    return res.redirect('/admin/login');
  }

  const cleanEmail = email ? email.trim().toLowerCase() : '';
  const cleanAdmin = adminEmail.trim().toLowerCase();
  
  const isEmailMatch = (cleanEmail === cleanAdmin) || 
                       (cleanEmail === cleanAdmin.replace('@gmail.com', '')) ||
                       (cleanEmail + '@gmail.com' === cleanAdmin);

  if (isEmailMatch && password === adminPassword) {
    req.session.admin = { email: adminEmail, name: 'Admin' };
    res.redirect('/admin/dashboard');
  } else {
    req.flash('error', 'Invalid admin credentials.');
    res.redirect('/admin/login');
  }
});

// Admin Logout
router.get('/admin/logout', (req, res) => {
  req.session.admin = null;
  res.redirect('/admin/login');
});

// Admin Dashboard (GET)
router.get('/admin/dashboard', requireAdmin, async (req, res) => {
  try {
    const users = await excel.readData('users');
    const bookings = await excel.readData('bookings');
    const reviews = await excel.readData('reviews');
    const contacts = await excel.readData('contacts');
    const newsletter = await excel.readData('newsletter');
    const settings = await excel.readData('settings');
    const team = await excel.readData('team');
    
    // Find settings variables
    const phoneSetting = settings.find(s => s.key === 'contact_phone');
    const emailSetting = settings.find(s => s.key === 'contact_email');
    const currentPhone = phoneSetting ? phoneSetting.value : '+91 98765 43210';
    const currentEmail = emailSetting ? emailSetting.value : 'xyz7@gmail.com';

    res.render('admin-dashboard', {
      title: 'Admin Dashboard - Ganesh Travels',
      admin: req.session.admin,
      users, bookings, reviews, contacts, newsletter, settings, team,
      currentPhone, currentEmail,
      messages: req.flash()
    });
  } catch (err) {
    console.error(err);
    req.flash('error', 'Failed to load dashboard data.');
    res.redirect('/admin/login');
  }
});

// Admin Update Agency Contact Settings (Phone & Email)
router.post('/admin/settings/edit', requireAdmin, async (req, res) => {
  try {
    const { contact_phone, contact_email } = req.body;
    if (!contact_phone || !contact_email) {
      req.flash('error', 'Phone number and Email are required.');
      return res.redirect('/admin/dashboard');
    }
    
    const phoneRow = await excel.findOne('settings', 'key', 'contact_phone');
    if (phoneRow) {
      await excel.updateRow('settings', 'contact_phone', { value: contact_phone });
    } else {
      await excel.addRow('settings', { key: 'contact_phone', value: contact_phone, createdAt: new Date().toLocaleString() });
    }
    
    const emailRow = await excel.findOne('settings', 'key', 'contact_email');
    if (emailRow) {
      await excel.updateRow('settings', 'contact_email', { value: contact_email });
    } else {
      await excel.addRow('settings', { key: 'contact_email', value: contact_email, createdAt: new Date().toLocaleString() });
    }
    
    req.flash('success', 'Agency contact details updated successfully!');
    res.redirect('/admin/dashboard');
  } catch (err) {
    console.error(err);
    req.flash('error', 'Failed to update agency settings.');
    res.redirect('/admin/dashboard');
  }
});

// Admin Add Team Member
router.post('/admin/team/add', requireAdmin, async (req, res) => {
  try {
    const { name, role, image } = req.body;
    if (!name || !role) {
      req.flash('error', 'Name and Role are required.');
      return res.redirect('/admin/dashboard');
    }
    await excel.addRow('team', {
      id: uuidv4(),
      name,
      role,
      image: image || 'https://images.unsplash.com/photo-1507003211169-0a1dd7228f2d?w=300',
      createdAt: new Date().toLocaleString()
    });
    req.flash('success', 'Team member added successfully!');
    res.redirect('/admin/dashboard');
  } catch (err) {
    console.error(err);
    req.flash('error', 'Failed to add team member.');
    res.redirect('/admin/dashboard');
  }
});

// Admin Delete Record
router.post('/admin/delete/:sheet/:id', requireAdmin, async (req, res) => {
  try {
    const { sheet, id } = req.params;
    const validSheets = ['users', 'bookings', 'reviews', 'contacts', 'newsletter', 'settings', 'team'];
    if (!validSheets.includes(sheet)) {
      req.flash('error', 'Invalid sheet name.');
      return res.redirect('/admin/dashboard');
    }
    await excel.deleteRow(sheet, id);
    req.flash('success', `Record deleted from ${sheet} successfully.`);
    res.redirect('/admin/dashboard');
  } catch (err) {
    console.error(err);
    req.flash('error', 'Failed to delete record.');
    res.redirect('/admin/dashboard');
  }
});

// Admin Download Excel File
router.get('/admin/download/:sheet', requireAdmin, async (req, res) => {
  const { sheet } = req.params;
  const validSheets = ['users', 'bookings', 'reviews', 'contacts', 'newsletter', 'settings', 'team'];
  if (!validSheets.includes(sheet)) {
    req.flash('error', 'Invalid file.');
    return res.redirect('/admin/dashboard');
  }
  try {
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${sheet}.xlsx`);
    await excel.writeExcelToStream(sheet, res);
  } catch (err) {
    console.error(`[Excel Download Error] Failed to generate ${sheet}.xlsx:`, err);
    req.flash('error', 'Could not generate Excel file.');
    res.redirect('/admin/dashboard');
  }
});

module.exports = router;
