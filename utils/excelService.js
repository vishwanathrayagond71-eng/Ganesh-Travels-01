// ============================================================
// excelService.js - Hybrid Database (Excel / MongoDB) Service
// Handles transparent data operations for local Excel & MongoDB Atlas
// ============================================================

const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const { MongoClient } = require('mongodb');

// Path to excel-data folder
const DATA_DIR = (process.env.NETLIFY || process.env.LAMBDA_TASK_ROOT)
  ? '/tmp/excel-data'
  : path.join(__dirname, '..', 'excel-data');

// Make sure excel-data folder exists
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

// Define headers for each Excel sheet / collection
const SHEET_HEADERS = {
  users: ['id', 'name', 'email', 'phone', 'password', 'createdAt'],
  bookings: ['id', 'userId', 'userName', 'userEmail', 'destination', 'date', 'guests', 'name', 'email', 'phone', 'message', 'status', 'guideName', 'createdAt'],
  reviews: ['id', 'userEmail', 'userName', 'destination', 'rating', 'comment', 'status', 'createdAt'],
  contacts: ['id', 'name', 'email', 'subject', 'message', 'createdAt'],
  newsletter: ['id', 'email', 'createdAt'],
  settings: ['key', 'value', 'createdAt'],
  team: ['id', 'name', 'role', 'image', 'createdAt'],
  pois: ['id', 'name', 'category', 'lat', 'lng', 'description', 'address', 'rating', 'image', 'createdAt'],
  destinations: ['id', 'name', 'category', 'country', 'image', 'description', 'price', 'rating', 'reviews', 'lat', 'lng', 'guideName', 'createdAt']
};

const MONGODB_URI = process.env.MONGODB_URI;
let mongoClient = null;
let db = null;
const isMongo = !!MONGODB_URI;

// Helper to get local excel file path
function getFilePath(sheetName) {
  return path.join(DATA_DIR, `${sheetName}.xlsx`);
}

// Initialize an excel file with headers if it doesn't exist
async function initSheet(sheetName, customWorkbook) {
  const filePath = getFilePath(sheetName);
  if (!fs.existsSync(filePath) || customWorkbook) {
    const workbook = customWorkbook || new ExcelJS.Workbook();
    let sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      sheet = workbook.addWorksheet(sheetName);
    }
    const headers = SHEET_HEADERS[sheetName];
    sheet.spliceRows(1, sheet.rowCount); // Clear sheet rows if rewriting
    sheet.addRow(headers);
    // Style the header row
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1d3557' }
    };
    if (!customWorkbook) {
      await workbook.xlsx.writeFile(filePath);
    }
  }
}

// Sync MongoDB collection to Excel file in the background (for Admin Panel downloads)
async function syncToExcel(sheetName) {
  if (!isMongo) return;
  try {
    const filePath = getFilePath(sheetName);
    const workbook = new ExcelJS.Workbook();
    await initSheet(sheetName, workbook);
    const sheet = workbook.getWorksheet(sheetName);
    const headers = SHEET_HEADERS[sheetName];
    
    const collection = db.collection(sheetName);
    const docs = await collection.find({}).toArray();
    
    docs.forEach(doc => {
      const rowData = headers.map(h => doc[h] || '');
      sheet.addRow(rowData);
    });
    
    await workbook.xlsx.writeFile(filePath);
  } catch (err) {
    console.error(`[Sync Error] Failed to sync ${sheetName} to Excel:`, err);
  }
}

// Initialize database / sheets
async function initAllSheets() {
  if (isMongo) {
    console.log('🔌 Connecting to MongoDB Atlas...');
    try {
      mongoClient = new MongoClient(MONGODB_URI);
      await mongoClient.connect();
      db = mongoClient.db();
      console.log(`✅ Connected successfully to MongoDB database: "${db.databaseName}"`);

      // Seed settings if empty
      const phone = await findOne('settings', 'key', 'contact_phone');
      if (!phone) {
        await addRow('settings', { key: 'contact_phone', value: '+91 98765 43210', createdAt: new Date().toLocaleString() });
      }
      const email = await findOne('settings', 'key', 'contact_email');
      if (!email) {
        await addRow('settings', { key: 'contact_email', value: 'xyz7@gmail.com', createdAt: new Date().toLocaleString() });
      }

      // Start background sync of all collections to local Excel files for download support
      for (const sheetName of Object.keys(SHEET_HEADERS)) {
        syncToExcel(sheetName);
      }
    } catch (err) {
      console.error('❌ Failed to connect to MongoDB:', err);
      throw err;
    }
  } else {
    // Fallback Excel initialization
    for (const sheetName of Object.keys(SHEET_HEADERS)) {
      await initSheet(sheetName);
    }

    const phone = await findOne('settings', 'key', 'contact_phone');
    if (!phone) {
      await addRow('settings', { key: 'contact_phone', value: '+91 98765 43210', createdAt: new Date().toLocaleString() });
    }
    const email = await findOne('settings', 'key', 'contact_email');
    if (!email) {
      await addRow('settings', { key: 'contact_email', value: 'xyz7@gmail.com', createdAt: new Date().toLocaleString() });
    }
    console.log('✅ Excel sheets initialized successfully (Local Mode)');
  }

  // Seed default POIs if empty
  try {
    const pois = await readData('pois');
    if (!pois || pois.length === 0) {
      const defaultPois = [
        {
          id: 'poi-1',
          name: 'Basaveshwar Temple, Vidyagiri',
          category: 'temple',
          lat: '16.1812',
          lng: '75.6983',
          description: 'A serene and sacred temple dedicated to Lord Basaveshwara, situated in Vidyagiri, Bagalkote.',
          address: 'Vidyagiri, Bagalkote, Karnataka',
          rating: '4.8',
          image: 'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-2',
          name: 'Badami Cave Temples',
          category: 'tourist',
          lat: '15.9189',
          lng: '75.6791',
          description: 'Magnificent rock-cut temples carved out of sandstone hills, dating back to the Chalukya dynasty in the 6th century.',
          address: 'Badami, Bagalkote District, Karnataka',
          rating: '4.9',
          image: 'https://images.unsplash.com/photo-1602216056096-3b40cc0c9944?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-3',
          name: 'Pattadakal Temple Complex',
          category: 'temple',
          lat: '15.9491',
          lng: '75.8197',
          description: 'A UNESCO World Heritage Site featuring a harmonious blend of North and South Indian temple architectures.',
          address: 'Pattadakal, Bagalkote District, Karnataka',
          rating: '4.8',
          image: 'https://images.unsplash.com/photo-1564507592333-c60657eea523?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-4',
          name: 'Aihole Durga Temple',
          category: 'tourist',
          lat: '16.0211',
          lng: '75.8828',
          description: 'Famous as the "Cradle of Indian Temple Architecture", Aihole features a unique Durga temple.',
          address: 'Aihole, Bagalkote District, Karnataka',
          rating: '4.7',
          image: 'https://images.unsplash.com/photo-1599661046289-e31897846e41?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-5',
          name: 'Vidyagiri Food Junction',
          category: 'restaurant',
          lat: '16.1834',
          lng: '75.6961',
          description: 'A popular local dining spot offering authentic North Karnataka meals, jolada rotti, and continental options.',
          address: 'Main Road, Vidyagiri, Bagalkote, Karnataka',
          rating: '4.5',
          image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-6',
          name: 'Almatti Dam & Rock Gardens',
          category: 'tourist',
          lat: '16.3323',
          lng: '75.8890',
          description: 'A major hydroelectric project on the Krishna River, featuring beautifully landscaped gardens and musical fountains.',
          address: 'Almatti, Bagalkote District, Karnataka',
          rating: '4.6',
          image: 'https://images.unsplash.com/photo-1469474968028-56623f02e42e?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-7',
          name: 'Taj Mahal, Agra',
          category: 'tourist',
          lat: '27.1751',
          lng: '78.0421',
          description: 'One of the Seven Wonders of the World — a magnificent ivory-white marble mausoleum on the south bank of Yamuna river.',
          address: 'Agra, Uttar Pradesh',
          rating: '4.9',
          image: 'https://images.unsplash.com/photo-1564507592333-c60657eea523?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-8',
          name: 'Grand Kaveri Veg Restaurant',
          category: 'restaurant',
          lat: '16.1805',
          lng: '75.6970',
          description: 'Excellent multi-cuisine family restaurant specializing in South Indian breakfast and authentic North Karnataka thalis.',
          address: 'Sector 10, Vidyagiri, Bagalkote, Karnataka',
          rating: '4.6',
          image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-9',
          name: 'Spice Garden Bar & Restaurant',
          category: 'restaurant',
          lat: '16.1840',
          lng: '75.6990',
          description: 'A popular family eatery serving spicy Mughlai, Chinese, and Tandoori chicken starters in Vidyagiri.',
          address: 'Club Road, Vidyagiri, Bagalkote, Karnataka',
          rating: '4.3',
          image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-10',
          name: 'Kalyan Heritage Hotel & Lodge',
          category: 'lodge',
          lat: '16.1825',
          lng: '75.6945',
          description: 'Comfortable premium rooms, standard lodges, and guest suites with modern amenities, room service, and ample parking.',
          address: 'Vidyagiri Main Road, Bagalkote, Karnataka',
          rating: '4.5',
          image: 'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-11',
          name: 'Hotel BG Palace & Lodge',
          category: 'lodge',
          lat: '16.1850',
          lng: '75.7020',
          description: 'Budget-friendly AC and Non-AC lodging accommodation ideal for families and travelers visiting Bagalkote.',
          address: 'Near Navanagar Gate, Vidyagiri, Bagalkote, Karnataka',
          rating: '4.2',
          image: 'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-12',
          name: 'Heritage Lodge & Resort Badami',
          category: 'lodge',
          lat: '15.9120',
          lng: '75.6830',
          description: 'Relaxing resort style cottages and traditional suites, located within 1 km of the historic Badami cave temples.',
          address: 'Station Road, Badami, Karnataka',
          rating: '4.7',
          image: 'https://images.unsplash.com/photo-1602216056096-3b40cc0c9944?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-13',
          name: 'Gokak Waterfalls',
          category: 'falls',
          lat: '16.1894',
          lng: '74.8317',
          description: 'A beautiful waterfall resembling Niagara Falls, formed by the Ghataprabha River dropping 52 meters over sandstone cliffs.',
          address: 'Gokak, Belagavi District, Karnataka',
          rating: '4.8',
          image: 'https://images.unsplash.com/photo-1469474968028-56623f02e42e?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-14',
          name: 'Sogal Someshwar Waterfalls',
          category: 'falls',
          lat: '15.7833',
          lng: '75.0500',
          description: 'A scenic hill side waterfall cascading near the ancient Someshwar Temple, ideal for weekend nature trips.',
          address: 'Sogal, Belagavi District, Karnataka',
          rating: '4.5',
          image: 'https://images.unsplash.com/photo-1469474968028-56623f02e42e?w=600',
          createdAt: new Date().toLocaleString()
        },
        // --- NEW SEEDED POIS FOR AGRA ---
        {
          id: 'poi-15',
          name: 'Agra Fort',
          category: 'tourist',
          lat: '27.1795',
          lng: '78.0211',
          description: 'A historical fort in the city of Agra in India. It was the main residence of the emperors of the Mughal Dynasty.',
          address: 'Agra Fort, Rakabganj, Agra, Uttar Pradesh',
          rating: '4.8',
          image: 'https://images.unsplash.com/photo-1599661046289-e31897846e41?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-16',
          name: 'Mehtab Bagh',
          category: 'tourist',
          lat: '27.1799',
          lng: '78.0423',
          description: 'A charbagh garden complex located in Agra, lying north of the Taj Mahal complex on the opposite side of the Yamuna River.',
          address: 'Opposite Taj Mahal, Dharmapuri, Forest Colony, Agra, Uttar Pradesh',
          rating: '4.5',
          image: 'https://images.unsplash.com/photo-1564507592333-c60657eea523?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-17',
          name: 'Pinch of Spice Restaurant',
          category: 'restaurant',
          lat: '27.2012',
          lng: '78.0078',
          description: 'Agra\'s premier fine dining restaurant specializing in delicious North Indian, Mughlai, and Tandoori cuisines.',
          address: '107/1-A, Wazirpura Road, Sanjay Place, Civil Lines, Agra, Uttar Pradesh',
          rating: '4.6',
          image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-18',
          name: 'Radisson Hotel Agra',
          category: 'lodge',
          lat: '27.1610',
          lng: '78.0458',
          description: 'An upscale hotel offering spacious rooms, an outdoor pool, and multiple restaurants with direct views of the Taj Mahal.',
          address: 'C-1, C-2, Fatehabad Road, Taj Nagari Phase 1, Agra, Uttar Pradesh',
          rating: '4.5',
          image: 'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-19',
          name: 'Mankameshwar Temple',
          category: 'temple',
          lat: '27.1898',
          lng: '78.0125',
          description: 'An ancient temple dedicated to Lord Shiva, located in the heart of Agra near Agra Fort Railway Station.',
          address: 'Daresi Road, Rawatpara, Sheebazar, Mantola, Agra, Uttar Pradesh',
          rating: '4.7',
          image: 'https://images.unsplash.com/photo-1561361058-c24e01e34f44?w=600',
          createdAt: new Date().toLocaleString()
        },
        // --- NEW SEEDED POIS FOR GOA ---
        {
          id: 'poi-20',
          name: 'Calangute Beach',
          category: 'tourist',
          lat: '15.5442',
          lng: '73.7550',
          description: 'Known as the "Queen of Beaches", it is the largest and most popular beach in North Goa, offering parasailing and water sports.',
          address: 'Calangute, North Goa, Goa',
          rating: '4.6',
          image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-21',
          name: 'Basilica of Bom Jesus',
          category: 'temple',
          lat: '15.5009',
          lng: '73.9116',
          description: 'A UNESCO World Heritage Site holding the mortal remains of St. Francis Xavier, renowned for its Baroque architecture.',
          address: 'Old Goa Road, Bainguinim, Goa',
          rating: '4.8',
          image: 'https://images.unsplash.com/photo-1564507592333-c60657eea523?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-22',
          name: 'Curlies Beach Shack',
          category: 'restaurant',
          lat: '15.5782',
          lng: '73.7422',
          description: 'A famous beach shack at Anjuna Beach, offering stunning sea views, delicious seafood, and vibrant music.',
          address: 'Anjuna Beach, Monteiro Vaddo, Anjuna, Goa',
          rating: '4.4',
          image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-23',
          name: 'Santana Beach Resort',
          category: 'lodge',
          lat: '15.5262',
          lng: '73.7610',
          description: 'A beautiful resort with Portuguese-styled cottages and swimming pools, located just steps away from Candolim Beach.',
          address: 'Candolim Beach Road, Candolim, Goa',
          rating: '4.5',
          image: 'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-24',
          name: 'Dudhsagar Waterfalls',
          category: 'falls',
          lat: '15.3185',
          lng: '74.3140',
          description: 'A four-tiered waterfall on the Mandovi River, looking like a cascade of milk flowing down the Western Ghats.',
          address: 'Sonaulim, Goa / Karnataka Border',
          rating: '4.8',
          image: 'https://images.unsplash.com/photo-1469474968028-56623f02e42e?w=600',
          createdAt: new Date().toLocaleString()
        },
        // --- NEW SEEDED POIS FOR JAIPUR ---
        {
          id: 'poi-25',
          name: 'Hawa Mahal',
          category: 'tourist',
          lat: '26.9239',
          lng: '75.8267',
          description: 'The "Palace of Winds", a five-story pyramidal shaped monument built with red and pink sandstone with 953 small windows.',
          address: 'Hawa Mahal Road, Badi Choupad, J.D.A. Market, Jaipur, Rajasthan',
          rating: '4.7',
          image: 'https://images.unsplash.com/photo-1599661046289-e31897846e41?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-26',
          name: 'Amer Fort',
          category: 'tourist',
          lat: '26.9855',
          lng: '75.8513',
          description: 'A majestic fort located on a hill in Amer town, famous for its artistic style elements, large ramparts, and series of gates.',
          address: 'Devisinghpura, Amer, Jaipur, Rajasthan',
          rating: '4.9',
          image: 'https://images.unsplash.com/photo-1477587458883-47145ed68d6a?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-27',
          name: 'Birla Mandir',
          category: 'temple',
          lat: '26.8920',
          lng: '75.8150',
          description: 'A grand Hindu temple built entirely in high-quality white marble, dedicated to Lord Vishnu and Goddess Lakshmi.',
          address: 'Jawahar Lal Nehru Marg, Tilak Nagar, Jaipur, Rajasthan',
          rating: '4.8',
          image: 'https://images.unsplash.com/photo-1564507592333-c60657eea523?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-28',
          name: 'Chokhi Dhani Resort',
          category: 'restaurant',
          lat: '26.7667',
          lng: '75.8356',
          description: 'A traditional Rajasthani ethnic village resort offering cultural shows, camel rides, and authentic Rajasthani thali feasts.',
          address: '12 Mile, Tonk Road, Via Vatika, Jaipur, Rajasthan',
          rating: '4.6',
          image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-29',
          name: 'Rambagh Palace Hotel',
          category: 'lodge',
          lat: '26.8981',
          lng: '75.8098',
          description: 'A world-famous luxury heritage hotel, formerly the residence of the Maharaja of Jaipur, featuring opulent gardens and suites.',
          address: 'Bhawani Singh Road, Rambagh Crossing, Jaipur, Rajasthan',
          rating: '4.9',
          image: 'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=600',
          createdAt: new Date().toLocaleString()
        },
        // --- NEW SEEDED POIS FOR MANALI ---
        {
          id: 'poi-30',
          name: 'Hadimba Devi Temple',
          category: 'temple',
          lat: '32.2476',
          lng: '77.1798',
          description: 'An ancient wooden cave temple dedicated to Hidimbi Devi, surrounded by a gorgeous cedar forest in Manali.',
          address: 'Hadimba Temple Road, Old Manali, Himachal Pradesh',
          rating: '4.7',
          image: 'https://images.unsplash.com/photo-1626621341517-bbf3d9990a23?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-31',
          name: 'Solang Valley Adventure Spot',
          category: 'tourist',
          lat: '32.3166',
          lng: '77.1500',
          description: 'A scenic side valley at the top of the Kullu Valley, famous for its summer and winter adventure sports like paragliding and skiing.',
          address: 'Solang, Manali, Himachal Pradesh',
          rating: '4.8',
          image: 'https://images.unsplash.com/photo-1626621341517-bbf3d9990a23?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-32',
          name: 'Jogini Waterfalls',
          category: 'falls',
          lat: '32.2677',
          lng: '77.1950',
          description: 'A beautiful waterfall cascading down rock cliffs, reachable via a scenic 2-kilometer trek from Vashisht Village.',
          address: 'Vashisht, Manali, Himachal Pradesh',
          rating: '4.6',
          image: 'https://images.unsplash.com/photo-1469474968028-56623f02e42e?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-33',
          name: 'Johnson Bar & Restaurant',
          category: 'restaurant',
          lat: '32.2452',
          lng: '77.1873',
          description: 'A popular restaurant with a cozy fireplace and garden seating, famous for its wood-fired pizzas and fresh trout dishes.',
          address: 'Circuit House Road, Manali, Himachal Pradesh',
          rating: '4.4',
          image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=600',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-34',
          name: 'The Himalayan Resort & Spa',
          category: 'lodge',
          lat: '32.2422',
          lng: '77.1923',
          description: 'Built in Victorian Gothic style, this premium castle-like resort offers luxury rooms, spa treatments, and stunning peak views.',
          address: 'Hadimba Road, Manali, Himachal Pradesh',
          rating: '4.7',
          image: 'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=600',
          createdAt: new Date().toLocaleString()
        },
        // --- NEW SEEDED INTERNATIONAL POIS FOR PARIS ---
        {
          id: 'poi-35',
          name: 'Eiffel Tower',
          category: 'tourist',
          lat: '48.8584',
          lng: '2.2945',
          description: 'The iconic wrought-iron lattice tower on the Champ de Mars in Paris, named after the engineer Gustave Eiffel.',
          address: 'Champ de Mars, 5 Avenue Anatole France, Paris, France',
          rating: '4.9',
          image: 'https://images.unsplash.com/photo-1502602898657-3e91760cbb34?w=300',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-36',
          name: 'Louvre Museum',
          category: 'tourist',
          lat: '48.8606',
          lng: '2.3376',
          description: 'The world\'s largest art museum and a historic monument in Paris, home to the Mona Lisa and Venus de Milo.',
          address: 'Rue de Rivoli, Paris, France',
          rating: '4.8',
          image: 'https://images.unsplash.com/photo-1502602898657-3e91760cbb34?w=300',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-37',
          name: 'Notre-Dame Cathedral',
          category: 'temple',
          lat: '48.8530',
          lng: '2.3499',
          description: 'A medieval Catholic cathedral on the Île de la Cité in the 4th arrondissement of Paris, renowned for Gothic architecture.',
          address: '6 Parvis Notre-Dame - Pl. Jean-Paul II, Paris, France',
          rating: '4.7',
          image: 'https://images.unsplash.com/photo-1502602898657-3e91760cbb34?w=300',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-38',
          name: 'Le Bistro Parisien',
          category: 'restaurant',
          lat: '48.8592',
          lng: '2.2980',
          description: 'A beautiful riverside bistro serving classic French cuisine with an unobstructed view of the Eiffel Tower.',
          address: 'Port de la Bourdonnais, Paris, France',
          rating: '4.3',
          image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=300',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-39',
          name: 'Hotel Regina Louvre',
          category: 'lodge',
          lat: '48.8631',
          lng: '2.3315',
          description: 'A premium 5-star luxury hotel in central Paris, offering elegant rooms and direct views of the Louvre Museum.',
          address: '2 Place des Pyramides, Paris, France',
          rating: '4.6',
          image: 'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=300',
          createdAt: new Date().toLocaleString()
        },
        // --- NEW SEEDED INTERNATIONAL POIS FOR DUBAI ---
        {
          id: 'poi-40',
          name: 'Burj Khalifa Tower',
          category: 'tourist',
          lat: '25.1972',
          lng: '55.2744',
          description: 'The tallest building in the world, standing at 828 meters, offering stunning 360-degree observation decks.',
          address: '1 Sheikh Mohammed bin Rashid Blvd, Downtown Dubai, Dubai, UAE',
          rating: '4.9',
          image: 'https://images.unsplash.com/photo-1512453979798-5ea266f8880c?w=300',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-41',
          name: 'Dubai Mall & Aquarium',
          category: 'tourist',
          lat: '25.1995',
          lng: '55.2796',
          description: 'The world\'s largest shopping and entertainment destination, featuring the Dubai Aquarium and Underwater Zoo.',
          address: 'Financial Center Rd, Downtown Dubai, Dubai, UAE',
          rating: '4.8',
          image: 'https://images.unsplash.com/photo-1512453979798-5ea266f8880c?w=300',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-42',
          name: 'Al Farooq Mosque',
          category: 'temple',
          lat: '25.1706',
          lng: '55.2289',
          description: 'An architectural marvel also known as the Blue Mosque, inspired by the historic Blue Mosque in Istanbul.',
          address: 'Al Safa 1, Dubai, UAE',
          rating: '4.8',
          image: 'https://images.unsplash.com/photo-1561361058-c24e01e34f44?w=300',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-43',
          name: 'Armani Restaurant Dubai',
          category: 'restaurant',
          lat: '25.1970',
          lng: '55.2742',
          description: 'Fine dining restaurant located in Burj Khalifa, serving award-winning Italian cuisine and chef selections.',
          address: 'Armani Hotel Dubai, Burj Khalifa, Dubai, UAE',
          rating: '4.6',
          image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=300',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-44',
          name: 'Address Downtown Hotel',
          category: 'lodge',
          lat: '25.1939',
          lng: '55.2818',
          description: 'A 5-star flagship hotel in Dubai overlooking Burj Khalifa and the Dubai Fountain, offering luxury amenities.',
          address: 'Sheikh Mohammed bin Rashid Blvd, Downtown Dubai, Dubai, UAE',
          rating: '4.7',
          image: 'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=300',
          createdAt: new Date().toLocaleString()
        },
        // --- NEW SEEDED INTERNATIONAL POIS FOR SINGAPORE ---
        {
          id: 'poi-45',
          name: 'Marina Bay Sands',
          category: 'tourist',
          lat: '1.2829',
          lng: '103.8587',
          description: 'The iconic resort featuring the world\'s largest rooftop infinity pool, a casino, and luxury shopping.',
          address: '10 Bayfront Ave, Singapore',
          rating: '4.9',
          image: 'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=300',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-46',
          name: 'Gardens by the Bay',
          category: 'tourist',
          lat: '1.2816',
          lng: '103.8636',
          description: 'A horticultural sanctuary featuring futuristic Supertree structures and the climate-controlled Flower Dome conservatory.',
          address: '18 Marina Gardens Dr, Singapore',
          rating: '4.8',
          image: 'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=300',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-47',
          name: 'Sri Mariamman Temple',
          category: 'temple',
          lat: '1.2826',
          lng: '103.8453',
          description: 'Singapore\'s oldest Hindu temple, featuring a majestic six-tiered gopuram adorned with colorful sculptures.',
          address: '244 South Bridge Rd, Chinatown, Singapore',
          rating: '4.7',
          image: 'https://images.unsplash.com/photo-1564507592333-c60657eea523?w=300',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-48',
          name: 'Jumbo Seafood Restaurant',
          category: 'restaurant',
          lat: '1.2891',
          lng: '103.8477',
          description: 'Famous dining spot on Clarke Quay, celebrated for serving Singapore\'s legendary Chilli Crab and black pepper crab.',
          address: '30 Merchant Rd, #01-01/02 Riverside Point, Singapore',
          rating: '4.5',
          image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=300',
          createdAt: new Date().toLocaleString()
        },
        {
          id: 'poi-49',
          name: 'Marina Bay Sands Luxury Lodge',
          category: 'lodge',
          lat: '1.2839',
          lng: '103.8596',
          description: 'Exclusive 5-star lodging suites on the upper decks of Marina Bay Sands, offering premium butler service.',
          address: '10 Bayfront Ave, Singapore',
          rating: '4.9',
          image: 'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=300',
          createdAt: new Date().toLocaleString()
        }
      ];
      for (const p of defaultPois) {
        await addRow('pois', p);
      }
      console.log('✅ Seeded default points of interest (POIs) successfully.');
    }
  } catch (err) {
    console.error('Failed to seed default POIs:', err);
  }

  // Seed default destinations if empty
  try {
    const dests = await readData('destinations');
    if (!dests || dests.length === 0) {
      const defaultDests = [
        { id: '1', name: 'Taj Mahal, Agra', category: 'historical', country: 'India', image: 'https://images.unsplash.com/photo-1564507592333-c60657eea523?w=600', description: 'One of the Seven Wonders of the World — a magnificent marble marble mausoleum in Agra.', price: '8999', rating: '4.9', reviews: '1240', lat: '27.1751', lng: '78.0421', guideName: 'Aditya Kumar', createdAt: new Date().toLocaleString() },
        { id: '2', name: 'Goa Beaches', category: 'beach', country: 'India', image: 'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=600', description: 'Sun-kissed beaches, vibrant nightlife, and Portuguese heritage await you in Goa.', price: '12999', rating: '4.7', reviews: '890', lat: '15.4989', lng: '73.8278', guideName: 'Rajesh G.', createdAt: new Date().toLocaleString() },
        { id: '3', name: 'Manali, Himachal', category: 'hill', country: 'India', image: 'https://images.unsplash.com/photo-1626621341517-bbf3d9990a23?w=600', description: 'Scenic hill station surrounded by snow-capped Himalayas, perfect for adventure seekers.', price: '14999', rating: '4.8', reviews: '760', lat: '32.2396', lng: '77.1887', guideName: 'Vikram Singh', createdAt: new Date().toLocaleString() },
        { id: '4', name: 'Jaipur, Rajasthan', category: 'historical', country: 'India', image: 'https://images.unsplash.com/photo-1599661046289-e31897846e41?w=600', description: 'The Pink City with majestic forts, palaces and vibrant bazaars of royal Rajasthan.', price: '9999', rating: '4.6', reviews: '654', lat: '26.9124', lng: '75.7873', guideName: 'Sanjay Sharma', createdAt: new Date().toLocaleString() },
        { id: '5', name: 'Kerala Backwaters', category: 'nature', country: 'India', image: 'https://images.unsplash.com/photo-1602216056096-3b40cc0c9944?w=600', description: 'Serene houseboat rides through lush green backwaters and coconut-fringed shores.', price: '16999', rating: '4.9', reviews: '920', lat: '9.4981', lng: '76.3388', guideName: 'Hari Prasad', createdAt: new Date().toLocaleString() },
        { id: '6', name: 'Varanasi, UP', category: 'spiritual', country: 'India', image: 'https://images.unsplash.com/photo-1561361058-c24e01e34f44?w=600', description: 'The spiritual capital of India — ancient ghats, sacred rituals, and divine Ganga Aarti.', price: '7999', rating: '4.5', reviews: '540', lat: '25.3176', lng: '82.9739', guideName: 'Ramesh Shastri', createdAt: new Date().toLocaleString() },
        { id: '7', name: 'Ladakh, J&K', category: 'hill', country: 'India', image: 'https://images.unsplash.com/photo-1536867114-e8b98cca2473?w=600', description: 'High-altitude desert landscape with monasteries, azure lakes and rugged mountains.', price: '24999', rating: '4.9', reviews: '430', lat: '34.1526', lng: '77.5771', guideName: 'Tashi Namgyal', createdAt: new Date().toLocaleString() },
        { id: '8', name: 'Andaman Islands', category: 'beach', country: 'India', image: 'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=600', description: 'Crystal-clear waters, pristine white-sand beaches and vibrant coral reefs.', price: '29999', rating: '4.8', reviews: '380', lat: '11.6234', lng: '92.7265', guideName: 'Subodh Das', createdAt: new Date().toLocaleString() },
        { id: '13', name: 'Bagalkote Heritage Hub', category: 'historical', country: 'India', image: 'https://images.unsplash.com/photo-1602216056096-3b40cc0c9944?w=600', description: 'Explore ancient Chalukyan rock-cut caves and temple architecture at Badami, Pattadakal, and Aihole.', price: '6999', rating: '4.8', reviews: '310', lat: '16.1812', lng: '75.6983', guideName: 'Ganesh Travels', createdAt: new Date().toLocaleString() },
      ];
      for (const d of defaultDests) {
        await addRow('destinations', d);
      }
      console.log('✅ Seeded default destinations successfully.');
    }

    // Ensure all famous Indian destinations are present (even if sheet already has some data)
    const currentDests = await readData('destinations');
    const extraIndianDests = [
      { name: 'Mumbai, Maharashtra', category: 'historical', country: 'India', image: 'https://images.unsplash.com/photo-1570168007204-dfb528c6958f?w=600', description: 'The financial capital of India, famous for the Gateway of India, Marine Drive, and Bollywood.', price: '7999', rating: '4.6', reviews: '450', lat: '19.0760', lng: '72.8777', guideName: 'Rahul Mehta' },
      { name: 'Delhi NCR', category: 'historical', country: 'India', image: 'https://images.unsplash.com/photo-1587474260584-136574528ed5?w=600', description: 'The historic capital of India, home to the Red Fort, Qutub Minar, and India Gate.', price: '6999', rating: '4.7', reviews: '610', lat: '28.6139', lng: '77.2090', guideName: 'Aman Verma' },
      { name: 'Bengaluru, Karnataka', category: 'nature', country: 'India', image: 'https://images.unsplash.com/photo-1596176530529-78163a4f7af2?w=600', description: 'The Silicon Valley of India, known for its beautiful parks, gardens, and pleasant climate.', price: '5999', rating: '4.5', reviews: '380', lat: '12.9716', lng: '77.5946', guideName: 'Kiran Kumar' },
      { name: 'Mysore, Karnataka', category: 'historical', country: 'India', image: 'https://images.unsplash.com/photo-1600100397608-f010e42ed987?w=600', description: 'The City of Palaces, famous for the magnificent Mysore Palace and vibrant Dasara festival.', price: '6499', rating: '4.8', reviews: '320', lat: '12.2958', lng: '76.6394', guideName: 'Guru Prasad' },
      { name: 'Hampi, Karnataka', category: 'historical', country: 'India', image: 'https://images.unsplash.com/photo-1600100398055-14f9764b3c76?w=600', description: 'A UNESCO World Heritage Site featuring the ruins of the historic Vijayanagara Empire.', price: '7499', rating: '4.9', reviews: '290', lat: '15.3350', lng: '76.4600', guideName: 'Shankar Gowda' },
      { name: 'Ooty, Tamil Nadu', category: 'hill', country: 'India', image: 'https://images.unsplash.com/photo-1590050752117-238cb0fb12b1?w=600', description: 'The Queen of Hill Stations, featuring beautiful tea gardens, lakes, and toy train rides.', price: '8999', rating: '4.7', reviews: '410', lat: '11.4102', lng: '76.6950', guideName: 'Ravi Pillai' },
      { name: 'Munnar, Kerala', category: 'hill', country: 'India', image: 'https://images.unsplash.com/photo-1593693397690-362cb9666fc2?w=600', description: 'Lush green tea plantations, misty mountains, and winding trails in Kerala.', price: '9499', rating: '4.8', reviews: '340', lat: '10.0889', lng: '77.0595', guideName: 'Jayan K.' },
      { name: 'Udaipur, Rajasthan', category: 'historical', country: 'India', image: 'https://images.unsplash.com/photo-1590050752117-238cb0fb12b1?w=600', description: 'The City of Lakes, famous for royal palaces, heritage boat rides, and romantic views.', price: '9999', rating: '4.8', reviews: '480', lat: '24.5854', lng: '73.7125', guideName: 'Dilip Singh' },
      { name: 'Shimla, Himachal', category: 'hill', country: 'India', image: 'https://images.unsplash.com/photo-1562690868-60bbe7293e94?w=600', description: 'The summer capital of British India, famous for Mall Road, colonial buildings, and snow.', price: '10999', rating: '4.6', reviews: '310', lat: '31.1048', lng: '77.1734', guideName: 'Devender Negi' },
      { name: 'Srinagar, J&K', category: 'hill', country: 'India', image: 'https://images.unsplash.com/photo-1566228015668-4c45dbc4e2f5?w=600', description: 'Paradise on Earth, famous for Dal Lake houseboats, shikara rides, and Tulip gardens.', price: '13999', rating: '4.9', reviews: '510', lat: '34.0837', lng: '74.7973', guideName: 'Showkat Mir' },
      { name: 'Amritsar, Punjab', category: 'spiritual', country: 'India', image: 'https://images.unsplash.com/photo-1514222134-b57cbb8ce073?w=600', description: 'Home of the Golden Temple, the holiest shrine of Sikhism, and Wagah Border.', price: '5999', rating: '4.9', reviews: '680', lat: '31.6340', lng: '74.8723', guideName: 'Jaspreet Singh' },
      { name: 'Rishikesh, Uttarakhand', category: 'spiritual', country: 'India', image: 'https://images.unsplash.com/photo-1545205597-3d9d02c29597?w=600', description: 'The Yoga Capital of the World, offering spiritual retreats and Ganga river rafting.', price: '6999', rating: '4.8', reviews: '430', lat: '30.0869', lng: '78.2676', guideName: 'Swami Dev' },
      { name: 'Darjeeling, West Bengal', category: 'hill', country: 'India', image: 'https://images.unsplash.com/photo-1557997637-bf482df97637?w=600', description: 'World-famous for its tea gardens, views of Mt. Kanchenjunga, and the Himalayan Toy Train.', price: '8499', rating: '4.7', reviews: '290', lat: '27.0410', lng: '88.2627', guideName: 'Pradhan Tamang' },
      { name: 'Pondicherry', category: 'beach', country: 'India', image: 'https://images.unsplash.com/photo-1582510003544-4d00b7f74220?w=600', description: 'A French colonial settlement, known for its beautiful beaches, heritage villas, and Auroville.', price: '7499', rating: '4.6', reviews: '320', lat: '11.9416', lng: '79.8083', guideName: 'Pierre M.' },
      { name: 'Coorg, Karnataka', category: 'nature', country: 'India', image: 'https://images.unsplash.com/photo-1590050752117-238cb0fb12b1?w=600', description: 'The Scotland of India, famous for coffee plantations, mist-covered hills, and waterfalls.', price: '7999', rating: '4.7', reviews: '250', lat: '12.4244', lng: '75.7382', guideName: 'Madappa K.' }
    ];

    let maxId = 0;
    currentDests.forEach(d => {
      const parsedId = parseInt(d.id);
      if (parsedId > maxId) maxId = parsedId;
    });

    for (const d of extraIndianDests) {
      const exists = currentDests.some(cd => cd.name.toLowerCase().includes(d.name.split(',')[0].toLowerCase()));
      if (!exists) {
        maxId++;
        const newDest = {
          id: String(maxId),
          name: d.name,
          category: d.category,
          country: d.country,
          image: d.image,
          description: d.description,
          price: d.price,
          rating: d.rating,
          reviews: d.reviews,
          lat: d.lat,
          lng: d.lng,
          guideName: d.guideName,
          createdAt: new Date().toLocaleString()
        };
        await addRow('destinations', newDest);
        console.log(`✅ Dynamically added famous Indian destination: ${d.name}`);
      }
    }
  } catch (err) {
    console.error('Failed to seed destinations:', err);
  }
}

// Read all rows (returns array of objects)
async function readData(sheetName) {
  if (isMongo) {
    try {
      const collection = db.collection(sheetName);
      const docs = await collection.find({}).toArray();
      return docs.map(doc => {
        const obj = { ...doc };
        delete obj._id; // Remove MongoDB internal ID for application consistency
        return obj;
      });
    } catch (err) {
      console.error(`[DB Error] Failed to read from collection ${sheetName}:`, err);
      return [];
    }
  } else {
    await initSheet(sheetName);
    const filePath = getFilePath(sheetName);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(sheetName);
    
    const rows = [];
    const headers = SHEET_HEADERS[sheetName];
    
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row
      const obj = {};
      headers.forEach((header, i) => {
        obj[header] = row.getCell(i + 1).value || '';
      });
      rows.push(obj);
    });
    
    return rows;
  }
}

// Add a new row / document
async function addRow(sheetName, data) {
  if (isMongo) {
    try {
      const collection = db.collection(sheetName);
      await collection.insertOne({ ...data });
      // Sync to local Excel in background for backup/download
      syncToExcel(sheetName);
    } catch (err) {
      console.error(`[DB Error] Failed to insert into collection ${sheetName}:`, err);
      throw err;
    }
  } else {
    await initSheet(sheetName);
    const filePath = getFilePath(sheetName);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(sheetName);
    
    const headers = SHEET_HEADERS[sheetName];
    const rowData = headers.map(h => data[h] || '');
    sheet.addRow(rowData);
    
    await workbook.xlsx.writeFile(filePath);
  }
}

// Delete a row / document by ID
async function deleteRow(sheetName, id) {
  if (isMongo) {
    try {
      const collection = db.collection(sheetName);
      const query = (sheetName === 'settings') ? { key: id } : { id: id };
      const result = await collection.deleteOne(query);
      syncToExcel(sheetName);
      return result.deletedCount > 0;
    } catch (err) {
      console.error(`[DB Error] Failed to delete from collection ${sheetName}:`, err);
      return false;
    }
  } else {
    await initSheet(sheetName);
    const filePath = getFilePath(sheetName);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(sheetName);
    
    let rowIndexToDelete = -1;
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      if (String(row.getCell(1).value) === String(id)) {
        rowIndexToDelete = rowNumber;
      }
    });
    
    if (rowIndexToDelete > 0) {
      sheet.spliceRows(rowIndexToDelete, 1);
      await workbook.xlsx.writeFile(filePath);
      return true;
    }
    return false;
  }
}

// Update a row / document by ID
async function updateRow(sheetName, id, newData) {
  if (isMongo) {
    try {
      const collection = db.collection(sheetName);
      const query = (sheetName === 'settings') ? { key: id } : { id: id };
      await collection.updateOne(query, { $set: newData });
      syncToExcel(sheetName);
    } catch (err) {
      console.error(`[DB Error] Failed to update collection ${sheetName}:`, err);
      throw err;
    }
  } else {
    await initSheet(sheetName);
    const filePath = getFilePath(sheetName);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(sheetName);
    const headers = SHEET_HEADERS[sheetName];
    
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      if (String(row.getCell(1).value) === String(id)) {
        headers.forEach((header, i) => {
          if (newData[header] !== undefined) {
            row.getCell(i + 1).value = newData[header];
          }
        });
      }
    });
    
    await workbook.xlsx.writeFile(filePath);
  }
}

// Find a single row / document by a field value
async function findOne(sheetName, field, value) {
  if (isMongo) {
    try {
      const collection = db.collection(sheetName);
      // Escape special regex characters in search term for safety
      const escapedValue = String(value).replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
      // Case-insensitive exact match
      const query = { [field]: { $regex: new RegExp(`^${escapedValue}$`, 'i') } };
      const doc = await collection.findOne(query);
      if (doc) {
        const obj = { ...doc };
        delete obj._id;
        return obj;
      }
      return null;
    } catch (err) {
      console.error(`[DB Error] Failed to find in collection ${sheetName}:`, err);
      return null;
    }
  } else {
    const rows = await readData(sheetName);
    return rows.find(row => String(row[field]).toLowerCase() === String(value).toLowerCase()) || null;
  }
}

// Generate Excel workbook in-memory and write to stream
async function writeExcelToStream(sheetName, stream) {
  const headers = SHEET_HEADERS[sheetName];
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(sheetName);
  
  sheet.addRow(headers);
  // Style the header row
  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1d3557' }
  };
  
  const data = await readData(sheetName);
  data.forEach(row => {
    const rowData = headers.map(h => row[h] !== undefined ? String(row[h]) : '');
    sheet.addRow(rowData);
  });
  
  await workbook.xlsx.write(stream);
}

module.exports = {
  initAllSheets,
  readData,
  addRow,
  deleteRow,
  updateRow,
  findOne,
  getFilePath,
  writeExcelToStream
};
