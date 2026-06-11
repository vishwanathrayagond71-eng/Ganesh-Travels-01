// ============================================================
// routes/index.js - All Application Routes
// Handles public pages, auth, user dashboard, admin panel
// ============================================================

const express = require('express');
const router = express.Router();
const bcrypt = require('bcryptjs');
const { v4: uuidv4 } = require('uuid');
const excel = require('../utils/excelService');
const osm = require('../utils/osmService');
const { requireUser, requireAdmin, redirectIfLoggedIn, redirectIfAdminLoggedIn } = require('../middleware/auth');
const path = require('path');

// Middleware moved to server.js globally

// ============================================================
// DATA: Destinations, Packages (Static data - no DB needed)
// ============================================================
// Helper to format destinations data types consistently
function formatDestinations(dests) {
  return dests.map(d => ({
    ...d,
    id: parseInt(d.id) || d.id,
    price: parseFloat(d.price) || 0,
    rating: parseFloat(d.rating) || 5.0,
    reviews: parseInt(d.reviews) || 0,
    lat: parseFloat(d.lat) || 0,
    lng: parseFloat(d.lng) || 0
  }));
}

const packages = [
  { id: 1, name: 'Golden Triangle Family Tour', type: 'family', image: 'https://images.unsplash.com/photo-1599661046289-e31897846e41?w=600', price: 34999, duration: '7 Days / 6 Nights', destinations: 'Delhi - Agra - Jaipur', services: ['Hotel (4★)', 'Breakfast & Dinner', 'AC Transport', 'Tour Guide', 'Entry Tickets'], rating: 4.8 },
  { id: 2, name: 'Kerala Honeymoon Special', type: 'honeymoon', image: 'https://images.unsplash.com/photo-1602216056096-3b40cc0c9944?w=600', price: 54999, duration: '6 Days / 5 Nights', destinations: 'Munnar - Thekkady - Alleppey', services: ['Luxury Hotel', 'All Meals', 'Houseboat Stay', 'Couple Spa', 'Romantic Dinner'], rating: 4.9 },
  { id: 3, name: 'Ladakh Adventure Expedition', type: 'adventure', image: 'https://images.unsplash.com/photo-1536867114-e8b98cca2473?w=600', price: 44999, duration: '9 Days / 8 Nights', destinations: 'Leh - Nubra - Pangong', services: ['Camp Stay', 'Breakfast', 'Bike Rental', 'Adventure Guide', 'Permits'], rating: 4.9 },
  { id: 4, name: 'Goa Solo Beach Getaway', type: 'solo', image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=600', price: 18999, duration: '4 Days / 3 Nights', destinations: 'North Goa - South Goa', services: ['Boutique Hotel', 'Breakfast', 'Scooter Rental', 'Beach Activities', 'City Tour'], rating: 4.6 },
  { id: 5, name: 'Rajasthan Group Heritage Tour', type: 'group', image: 'https://images.unsplash.com/photo-1477587458883-47145ed68d6a?w=600', price: 28999, duration: '8 Days / 7 Nights', destinations: 'Jaipur - Jodhpur - Udaipur', services: ['Heritage Hotel', 'Daily Breakfast', 'AC Bus', 'Camel Safari', 'Folk Show'], rating: 4.7 },
  { id: 6, name: 'Kashmir Paradise Group Tour', type: 'group', image: 'https://images.unsplash.com/photo-1566228015668-4c45dbc4e2f5?w=600', price: 39999, duration: '6 Days / 5 Nights', destinations: 'Srinagar - Gulmarg - Pahalgam', services: ['Luxury Hotel Stay', 'Daily Breakfast & Dinner', 'Houseboat Stay', 'Shikara Ride', 'Private AC Cab'], rating: 4.9 }
];

// ============================================================
// PUBLIC PAGES
// ============================================================

// Home Page
router.get('/', async (req, res) => {
  try {
    const reviews = await excel.readData('reviews');
    const approvedReviews = reviews.filter(r => r.status === 'approved').slice(-6);
    const dests = await excel.readData('destinations');
    res.render('home', {
      title: 'Ganesh Travels - Discover India & Beyond',
      user: req.session.user || null,
      destinations: formatDestinations(dests).slice(0, 6),
      packages: packages.slice(0, 3),
      reviews: approvedReviews,
      messages: req.flash()
    });
  } catch (err) {
    console.error(err);
    res.render('home', { title: 'Ganesh Travels', user: req.session.user || null, destinations: [], packages: packages.slice(0, 3), reviews: [], messages: {} });
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
router.get('/destinations', async (req, res) => {
  try {
    const dests = await excel.readData('destinations');
    res.render('destinations', { title: 'Destinations - Ganesh Travels', user: req.session.user || null, destinations: formatDestinations(dests), messages: req.flash() });
  } catch (err) {
    console.error(err);
    res.render('destinations', { title: 'Destinations - Ganesh Travels', user: req.session.user || null, destinations: [], messages: req.flash() });
  }
});

// Destination Detail Page with Map & Nearby POIs Guide
router.get('/destinations/:id', async (req, res) => {
  try {
    const dests = await excel.readData('destinations');
    const formattedDests = formatDestinations(dests);
    const destId = parseInt(req.params.id);
    const destination = formattedDests.find(d => d.id === destId);
    if (!destination) {
      req.flash('error', 'Destination not found.');
      return res.redirect('/destinations');
    }

    // Load POIs from Excel / MongoDB
    let pois = await excel.readData('pois');

    // Distance calculation helper (Haversine formula)
    const calculateDistance = (lat1, lon1, lat2, lon2) => {
      const R = 6371; // Earth radius in km
      const dLat = (lat2 - lat1) * Math.PI / 180;
      const dLon = (lon2 - lon1) * Math.PI / 180;
      const a = 
        Math.sin(dLat/2) * Math.sin(dLat/2) +
        Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * 
        Math.sin(dLon/2) * Math.sin(dLon/2);
      const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
      return R * c;
    };

    const destLat = parseFloat(destination.lat);
    const destLng = parseFloat(destination.lng);

    // Check if we need to fetch POIs from OSM (if < 5 POIs exist near the destination within 50km)
    let localPoisNearDest = pois.filter(poi => {
      const poiLat = parseFloat(poi.lat);
      const poiLng = parseFloat(poi.lng);
      if (isNaN(poiLat) || isNaN(poiLng)) return false;
      return calculateDistance(destLat, destLng, poiLat, poiLng) <= 50;
    });

    if (localPoisNearDest.length < 5) {
      console.log(`[Dynamic POI Seeding] Fetching POIs near ${destination.name} (${destLat}, ${destLng})...`);
      const cityName = destination.name.split(',')[0];
      try {
        const newPois = await osm.getPoisForCity(cityName, destLat, destLng);
        for (const poi of newPois) {
          const exists = pois.some(p => p.name.toLowerCase() === poi.name.toLowerCase() || 
            (Math.abs(parseFloat(p.lat) - parseFloat(poi.lat)) < 0.0001 && 
             Math.abs(parseFloat(p.lng) - parseFloat(poi.lng)) < 0.0001)
          );
          if (!exists) {
            await excel.addRow('pois', poi);
            pois.push(poi); // Add to in-memory list
          }
        }
      } catch (err) {
        console.error('[Dynamic POI Seeding Error]:', err.message);
      }
    }

    // Calculate distance and categorize POIs within 100km
    const nearbyPois = {
      tourist: [],
      temple: [],
      restaurant: [],
      lodge: [],
      falls: []
    };

    pois.forEach(poi => {
      const poiLat = parseFloat(poi.lat);
      const poiLng = parseFloat(poi.lng);
      if (isNaN(poiLat) || isNaN(poiLng)) return;

      const dist = calculateDistance(destLat, destLng, poiLat, poiLng);
      // Only include POIs within 100 km radius
      if (dist <= 100) {
        const poiWithDistance = { ...poi, distance: dist };
        let category = poi.category;
        if (category === 'hotel') category = 'lodge'; // Normalize hotel category to lodge
        
        if (nearbyPois[category]) {
          nearbyPois[category].push(poiWithDistance);
        } else {
          // If category matches none of the above, map it to tourist places
          nearbyPois.tourist.push(poiWithDistance);
        }
      }
    });

    // Sort nearby POIs by distance (closest first)
    for (const key in nearbyPois) {
      nearbyPois[key].sort((a, b) => a.distance - b.distance);
    }

    res.render('destination-detail', {
      title: `${destination.name} - Local Guide & Booking | Ganesh Travels`,
      user: req.session.user || null,
      destination,
      destinations: formattedDests, // passed to populate booking form options
      nearbyPois,
      messages: req.flash()
    });
  } catch (err) {
    console.error('[Destination Detail Page Error]:', err);
    req.flash('error', 'Unable to load destination details.');
    res.redirect('/destinations');
  }
});

// Explore GPS Map Page
router.get('/explore', async (req, res) => {
  try {
    const pois = await excel.readData('pois');
    res.render('explore', { title: 'Explore GPS - Ganesh Travels', user: req.session.user || null, pois, messages: req.flash() });
  } catch (err) {
    console.error('[Explore Route Error]:', err);
    res.render('explore', { title: 'Explore GPS - Ganesh Travels', user: req.session.user || null, pois: [], messages: req.flash() });
  }
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
router.get('/booking', async (req, res) => {
  try {
    const dests = await excel.readData('destinations');
    const preDestination = req.query.destination || '';
    res.render('booking', { title: 'Book Your Tour - Ganesh Travels', user: req.session.user || null, destinations: formatDestinations(dests), packages, preDestination, messages: req.flash() });
  } catch (err) {
    console.error(err);
    res.render('booking', { title: 'Book Your Tour - Ganesh Travels', user: req.session.user || null, destinations: [], packages, preDestination: '', messages: req.flash() });
  }
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
      name, email, phone, message: message || '', status: 'pending', guideName: '', createdAt: new Date().toLocaleString()
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
    const dests = await excel.readData('destinations');
    res.render('reviews', { title: 'Customer Reviews - Ganesh Travels', user: req.session.user || null, reviews: approvedReviews, destinations: formatDestinations(dests), messages: req.flash() });
  } catch (err) {
    res.render('reviews', { title: 'Reviews', user: req.session.user || null, reviews: [], destinations: [], messages: {} });
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
  const adminEmail = process.env.ADMIN_EMAIL || 'xyz7@gmail.com';
  const adminPassword = process.env.ADMIN_PASSWORD || '1234';

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
    const pois = await excel.readData('pois');
    const destinationsList = await excel.readData('destinations');
    
    // Find settings variables
    const phoneSetting = settings.find(s => s.key === 'contact_phone');
    const emailSetting = settings.find(s => s.key === 'contact_email');
    const currentPhone = phoneSetting ? phoneSetting.value : '+91 98765 43210';
    const currentEmail = emailSetting ? emailSetting.value : 'xyz7@gmail.com';

    res.render('admin-dashboard', {
      title: 'Admin Dashboard - Ganesh Travels',
      admin: req.session.admin,
      users, bookings, reviews, contacts, newsletter, settings, team, pois,
      destinations: formatDestinations(destinationsList),
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

// Admin Add Point of Interest (POI)
router.post('/admin/pois/add', requireAdmin, async (req, res) => {
  try {
    const { name, category, lat, lng, description, address, rating, image } = req.body;
    if (!name || !category || !lat || !lng) {
      req.flash('error', 'Name, Category, Latitude and Longitude are required.');
      return res.redirect('/admin/dashboard');
    }
    
    await excel.addRow('pois', {
      id: uuidv4(),
      name,
      category,
      lat: String(lat).trim(),
      lng: String(lng).trim(),
      description: description || '',
      address: address || '',
      rating: rating || '5.0',
      image: image || 'https://images.unsplash.com/photo-1469474968028-56623f02e42e?w=600',
      createdAt: new Date().toLocaleString()
    });
    
    req.flash('success', 'Point of Interest added successfully!');
    res.redirect('/admin/dashboard');
  } catch (err) {
    console.error('[Add POI Error]:', err);
    req.flash('error', 'Failed to add point of interest.');
    res.redirect('/admin/dashboard');
  }
});

// Admin Add Destination
router.post('/admin/destinations/add', requireAdmin, async (req, res) => {
  try {
    const { name, category, country, image, description, price, rating, reviews, lat, lng, guideName } = req.body;
    if (!name || !category || !country || !price || !lat || !lng) {
      req.flash('error', 'Required fields (Name, Category, Country, Price, Lat, Lng) are missing.');
      return res.redirect('/admin/dashboard');
    }
    
    // Find highest existing ID to increment it
    const dests = await excel.readData('destinations');
    let maxId = 0;
    dests.forEach(d => {
      const idNum = parseInt(d.id);
      if (idNum > maxId) maxId = idNum;
    });
    const newId = String(maxId + 1);

    await excel.addRow('destinations', {
      id: newId,
      name,
      category,
      country,
      image: image || 'https://images.unsplash.com/photo-1469474968028-56623f02e42e?w=600',
      description: description || '',
      price: String(price).trim(),
      rating: rating || '5.0',
      reviews: reviews || '0',
      lat: String(lat).trim(),
      lng: String(lng).trim(),
      guideName: guideName || '',
      createdAt: new Date().toLocaleString()
    });

    req.flash('success', 'Destination added successfully! 🗺️');
    res.redirect('/admin/dashboard');
  } catch (err) {
    console.error('[Add Destination Error]:', err);
    req.flash('error', 'Failed to add destination.');
    res.redirect('/admin/dashboard');
  }
});

// Admin Edit Destination Price and Guide
router.post('/admin/destinations/edit', requireAdmin, async (req, res) => {
  try {
    const { id, price, guideName } = req.body;
    if (!id || !price) {
      req.flash('error', 'Destination ID and Price are required to edit.');
      return res.redirect('/admin/dashboard');
    }

    await excel.updateRow('destinations', id, {
      price: String(price).trim(),
      guideName: guideName || ''
    });

    req.flash('success', 'Destination updated successfully! 💵');
    res.redirect('/admin/dashboard');
  } catch (err) {
    console.error('[Edit Destination Error]:', err);
    req.flash('error', 'Failed to update destination.');
    res.redirect('/admin/dashboard');
  }
});

// Admin Update Booking Status & Guide Allotment
router.post('/admin/bookings/update/:id', requireAdmin, async (req, res) => {
  try {
    const { id } = req.params;
    const { status, guideName } = req.body;
    if (!status) {
      req.flash('error', 'Status is required.');
      return res.redirect('/admin/dashboard');
    }

    await excel.updateRow('bookings', id, {
      status,
      guideName: guideName || ''
    });

    req.flash('success', 'Booking updated and guide allotted successfully! 💼');
    res.redirect('/admin/dashboard');
  } catch (err) {
    console.error('[Update Booking Error]:', err);
    req.flash('error', 'Failed to update booking.');
    res.redirect('/admin/dashboard');
  }
});

// Admin Delete Record
router.post('/admin/delete/:sheet/:id', requireAdmin, async (req, res) => {
  try {
    const { sheet, id } = req.params;
    const validSheets = ['users', 'bookings', 'reviews', 'contacts', 'newsletter', 'settings', 'team', 'pois', 'destinations'];
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
  const validSheets = ['users', 'bookings', 'reviews', 'contacts', 'newsletter', 'settings', 'team', 'pois', 'destinations'];
  if (!validSheets.includes(sheet)) {
    req.flash('error', 'Invalid file.');
    return res.redirect('/admin/dashboard');
  }
  try {
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${sheet}.xlsx`);
    await excel.writeExcelToStream(sheet, res);
    res.end();
  } catch (err) {
    console.error(`[Excel Download Error] Failed to generate ${sheet}.xlsx:`, err);
    req.flash('error', 'Could not generate Excel file.');
    res.redirect('/admin/dashboard');
  }
});

// Dynamic destinations search and geocoding API
router.get('/api/destinations/search', async (req, res) => {
  try {
    const q = req.query.q ? req.query.q.trim() : '';
    const dests = await excel.readData('destinations');
    const formattedDests = formatDestinations(dests);

    if (!q) {
      return res.json({ destinations: formattedDests, nearbyDestinations: [], nearbyPois: { tourist: [], temple: [], restaurant: [], lodge: [], falls: [] } });
    }

    const queryLower = q.toLowerCase();
    // 1. Search locally
    let matches = formattedDests.filter(d => 
      d.name.toLowerCase().includes(queryLower) || 
      d.description.toLowerCase().includes(queryLower) ||
      d.country.toLowerCase().includes(queryLower)
    );

    // 2. If 0 local matches, resolve using OpenStreetMap Nominatim
    if (matches.length === 0) {
      const resolved = await osm.geocodeAddress(q);
      if (resolved) {
        // Check if this city matches any local destination
        const existing = formattedDests.find(d => 
          (resolved.city && d.name.toLowerCase().includes(resolved.city.toLowerCase())) || 
          resolved.name.toLowerCase().includes(d.name.split(',')[0].toLowerCase())
        );

        if (existing) {
          matches = [existing];
        } else {
          // Determine category
          let category = 'historical';
          const nameLower = resolved.name.toLowerCase();
          if (nameLower.includes('beach') || nameLower.includes('sea') || nameLower.includes('coast') || nameLower.includes('island')) {
            category = 'beach';
          } else if (nameLower.includes('hill') || nameLower.includes('mountain') || nameLower.includes('peak') || nameLower.includes('valley') || nameLower.includes('alps') || nameLower.includes('himachal') || nameLower.includes('kashmir')) {
            category = 'hill';
          } else if (nameLower.includes('temple') || nameLower.includes('church') || nameLower.includes('mosque') || nameLower.includes('spiritual') || nameLower.includes('sacred') || nameLower.includes('shrine')) {
            category = 'spiritual';
          } else if (nameLower.includes('park') || nameLower.includes('falls') || nameLower.includes('lake') || nameLower.includes('river') || nameLower.includes('nature') || nameLower.includes('forest') || nameLower.includes('wildlife')) {
            category = 'nature';
          }

          // Create new destination ID (increment maximum)
          let maxId = 0;
          formattedDests.forEach(d => {
            const parsedId = parseInt(d.id);
            if (parsedId > maxId) maxId = parsedId;
          });
          const newId = String(maxId + 1);
          
          const cleanName = resolved.city ? `${resolved.city}, ${resolved.country}` : resolved.name.split(',').slice(0, 2).join(', ');
          
          const newDest = {
            id: newId,
            name: cleanName,
            category,
            country: resolved.country,
            image: 'https://images.unsplash.com/photo-1469474968028-56623f02e42e?w=600', // Default beautiful travel photo
            description: `A stunning and newly discovered destination in ${resolved.country}. Explore the beautiful scenery, rich cultural heritage, and local wonders of ${cleanName}.`,
            price: String(5000 + Math.floor(Math.random() * 8000)),
            rating: String((4.3 + Math.random() * 0.7).toFixed(1)),
            reviews: String(10 + Math.floor(Math.random() * 150)),
            lat: String(resolved.lat),
            lng: String(resolved.lng),
            guideName: 'Ganesh Travels Guide',
            createdAt: new Date().toLocaleString()
          };

          await excel.addRow('destinations', newDest);
          console.log(`✅ Dynamically added new destination from search: ${newDest.name}`);
          
          // Seed POIs for this new place in the background/synchronously
          try {
            const newPois = await osm.getPoisForCity(resolved.city || q, resolved.lat, resolved.lng);
            const currentPoisList = await excel.readData('pois');
            for (const poi of newPois) {
              const exists = currentPoisList.some(p => p.name.toLowerCase() === poi.name.toLowerCase());
              if (!exists) {
                await excel.addRow('pois', poi);
              }
            }
            console.log(`✅ Seeded ${newPois.length} POIs near ${newDest.name}`);
          } catch (poiErr) {
            console.error('[Dynamic POI Seeding from Search Error]:', poiErr.message);
          }

          const finalDest = {
            ...newDest,
            id: parseInt(newId),
            price: parseFloat(newDest.price),
            rating: parseFloat(newDest.rating),
            reviews: parseInt(newDest.reviews),
            lat: parseFloat(newDest.lat),
            lng: parseFloat(newDest.lng)
          };
          matches = [finalDest];
        }
      }
    }

    // 3. If we have a matched destination, calculate nearby destinations & POIs
    let nearbyDestinations = [];
    let nearbyPois = { tourist: [], temple: [], restaurant: [], lodge: [], falls: [] };

    // Distance calculation helper (Haversine formula)
    const calculateDistance = (lat1, lon1, lat2, lon2) => {
      const R = 6371; // Earth radius in km
      const dLat = (lat2 - lat1) * Math.PI / 180;
      const dLon = (lon2 - lon1) * Math.PI / 180;
      const a = 
        Math.sin(dLat/2) * Math.sin(dLat/2) +
        Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * 
        Math.sin(dLon/2) * Math.sin(dLon/2);
      const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
      return R * c;
    };

    if (matches.length > 0) {
      const target = matches[0];
      
      // Calculate nearby destinations (within 200km, excluding the target itself)
      nearbyDestinations = formattedDests
        .filter(d => d.id !== target.id)
        .map(d => ({
          ...d,
          distance: calculateDistance(target.lat, target.lng, d.lat, d.lng)
        }))
        .filter(d => d.distance <= 200)
        .sort((a, b) => a.distance - b.distance)
        .slice(0, 6);

      // Fetch nearby POIs (within 100km)
      const allPois = await excel.readData('pois');

      // Check if we need to fetch POIs from OSM (if < 5 POIs exist near the destination within 50km)
      let localPoisNearDest = allPois.filter(poi => {
        const poiLat = parseFloat(poi.lat);
        const poiLng = parseFloat(poi.lng);
        if (isNaN(poiLat) || isNaN(poiLng)) return false;
        return calculateDistance(target.lat, target.lng, poiLat, poiLng) <= 50;
      });

      if (localPoisNearDest.length < 5) {
        console.log(`[Dynamic POI Seeding from Search] Fetching POIs near ${target.name} (${target.lat}, ${target.lng})...`);
        const cityName = target.name.split(',')[0];
        try {
          const newPois = await osm.getPoisForCity(cityName, target.lat, target.lng);
          for (const poi of newPois) {
            const exists = allPois.some(p => p.name.toLowerCase() === poi.name.toLowerCase() || 
              (Math.abs(parseFloat(p.lat) - parseFloat(poi.lat)) < 0.0001 && 
               Math.abs(parseFloat(p.lng) - parseFloat(poi.lng)) < 0.0001)
            );
            if (!exists) {
              await excel.addRow('pois', poi);
              allPois.push(poi); // Add to local array so we can include it in response
            }
          }
        } catch (err) {
          console.error('[Dynamic POI Seeding from Search Error]:', err.message);
        }
      }

      allPois.forEach(poi => {
        const poiLat = parseFloat(poi.lat);
        const poiLng = parseFloat(poi.lng);
        if (isNaN(poiLat) || isNaN(poiLng)) return;

        const dist = calculateDistance(target.lat, target.lng, poiLat, poiLng);
        if (dist <= 100) {
          const poiWithDistance = { ...poi, distance: dist };
          let category = poi.category;
          if (category === 'hotel') category = 'lodge';

          if (nearbyPois[category]) {
            nearbyPois[category].push(poiWithDistance);
          } else {
            nearbyPois.tourist.push(poiWithDistance);
          }
        }
      });

      // Sort nearby POIs by distance
      for (const key in nearbyPois) {
        nearbyPois[key].sort((a, b) => a.distance - b.distance);
      }
    }

    res.json({
      destinations: matches,
      nearbyDestinations,
      nearbyPois
    });

  } catch (err) {
    console.error('[API Search Error]:', err);
    res.status(500).json({ error: 'Failed to search destinations' });
  }
});

// API to get nearby POIs (for Locate Me page)
router.get('/api/pois/nearby', async (req, res) => {
  try {
    const lat = parseFloat(req.query.lat);
    const lng = parseFloat(req.query.lng);
    if (isNaN(lat) || isNaN(lng)) {
      return res.status(400).json({ error: 'Invalid latitude or longitude parameters' });
    }

    const pois = await excel.readData('pois');

    const calculateDistance = (lat1, lon1, lat2, lon2) => {
      const R = 6371; // Earth radius in km
      const dLat = (lat2 - lat1) * Math.PI / 180;
      const dLon = (lon2 - lon1) * Math.PI / 180;
      const a = 
        Math.sin(dLat/2) * Math.sin(dLat/2) +
        Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * 
        Math.sin(dLon/2) * Math.sin(dLon/2);
      const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
      return R * c;
    };

    // Find POIs near user coordinates (within 50km)
    let nearby = pois.filter(poi => {
      const poiLat = parseFloat(poi.lat);
      const poiLng = parseFloat(poi.lng);
      if (isNaN(poiLat) || isNaN(poiLng)) return false;
      return calculateDistance(lat, lng, poiLat, poiLng) <= 50;
    });

    // If we have fewer than 8 local POIs within 50km, fetch real ones from OSM
    if (nearby.length < 8) {
      console.log(`[OSM GPS Seeding] Fetching new POIs near GPS coordinates ${lat}, ${lng}...`);
      try {
        const geoInfo = await osm.reverseGeocode(lat, lng);
        const searchCity = geoInfo ? geoInfo.city : `Location near ${lat.toFixed(2)},${lng.toFixed(2)}`;
        
        const newPois = await osm.getPoisForCity(searchCity, lat, lng);
        
        // Save new POIs to database if they don't already exist
        for (const poi of newPois) {
          const exists = pois.some(p => p.name.toLowerCase() === poi.name.toLowerCase() || 
            (Math.abs(parseFloat(p.lat) - parseFloat(poi.lat)) < 0.0001 && 
             Math.abs(parseFloat(p.lng) - parseFloat(poi.lng)) < 0.0001)
          );
          if (!exists) {
            await excel.addRow('pois', poi);
            pois.push(poi); // add to all list so we can include it in final output
          }
        }
      } catch (osmErr) {
        console.error('[OSM Locate Me Seeding Error]:', osmErr.message);
      }

      // Re-filter after adding the new ones, up to 50km radius so the map gets a good selection
      nearby = pois.filter(poi => {
        const poiLat = parseFloat(poi.lat);
        const poiLng = parseFloat(poi.lng);
        if (isNaN(poiLat) || isNaN(poiLng)) return false;
        return calculateDistance(lat, lng, poiLat, poiLng) <= 50;
      });
    }

    res.json(nearby);
  } catch (err) {
    console.error('[API Nearby POIs Error]:', err);
    res.status(500).json({ error: 'Failed to fetch nearby POIs' });
  }
});

module.exports = router;
