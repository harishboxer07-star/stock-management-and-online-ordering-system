const mongoose = require('mongoose');

const connectDB = async (retries = 3, delay = 2000) => {
  const primaryUri = process.env.MONGODB_URI || 'mongodb://127.0.0.1:27017/smart_grocery_db';
  const localUri = 'mongodb://127.0.0.1:27017/smart_grocery_db';

  // Try primary URI first
  try {
    const conn = await mongoose.connect(primaryUri, { serverSelectionTimeoutMS: 4000 });
    console.log(`[Database] MongoDB Connected: ${conn.connection.host}/${conn.connection.name}`);
    return conn;
  } catch (primaryErr) {
    console.warn(`[Database Warning] Primary connection failed: ${primaryErr.message}`);

    // If primary was a remote cluster and failed, automatically try local MongoDB
    if (primaryUri !== localUri && !primaryUri.includes('127.0.0.1') && !primaryUri.includes('localhost')) {
      try {
        console.log('[Database] Attempting fallback to local MongoDB (mongodb://127.0.0.1:27017/smart_grocery_db)...');
        const localConn = await mongoose.connect(localUri, { serverSelectionTimeoutMS: 4000 });
        console.log(`[Database] Connected to local MongoDB: ${localConn.connection.host}/${localConn.connection.name}`);
        return localConn;
      } catch (localErr) {
        console.warn(`[Database Warning] Local MongoDB fallback failed: ${localErr.message}`);
      }
    }
  }

  // If fallback was not applicable or failed, retry with primary
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      console.log(`[Database] Retrying connection (Attempt ${attempt}/${retries})...`);
      const conn = await mongoose.connect(primaryUri, { serverSelectionTimeoutMS: 4000 });
      console.log(`[Database] MongoDB Connected: ${conn.connection.host}/${conn.connection.name}`);
      return conn;
    } catch (error) {
      console.error(`[Database Connection Attempt ${attempt}/${retries}]: ${error.message}`);
      if (attempt === retries) {
        console.error('[Database Error] Max retries reached. Exiting...');
        process.exit(1);
      }
      console.log(`[Database] Retrying in ${delay / 1000}s...`);
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
  }
};

module.exports = connectDB;
