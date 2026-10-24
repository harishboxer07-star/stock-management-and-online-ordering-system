const mongoose = require('mongoose');
require('dotenv').config();

const User = require('../models/User');
const Product = require('../models/Product');
const Batch = require('../models/Batch');
const Supplier = require('../models/Supplier');
const Order = require('../models/Order');
const PurchaseOrder = require('../models/PurchaseOrder');
const DailySale = require('../models/DailySale');
const AuditLog = require('../models/AuditLog');

const seedDatabase = async () => {
  try {
    if (mongoose.connection.readyState !== 1) {
      await mongoose.connect(process.env.MONGODB_URI || 'mongodb://127.0.0.1:27017/smart_grocery_db');
    }
    console.log('[Seed] Connected to MongoDB. Clearing existing collections...');

    await User.deleteMany({});
    await Product.deleteMany({});
    await Batch.deleteMany({});
    await Supplier.deleteMany({});
    await Order.deleteMany({});
    await PurchaseOrder.deleteMany({});
    await DailySale.deleteMany({});
    await AuditLog.deleteMany({});

    console.log('[Seed] Creating demo users...');
    const adminUser = await User.create({
      name: 'Harish Kumar (Admin)',
      email: 'admin@grocery.com',
      password: 'admin123', // Will be hashed via pre-save hook
      role: 'admin',
      phone: '+91 9840123456',
      address: { street: '12 GST Road, Guindy', city: 'Chennai', state: 'Tamil Nadu', zipCode: '600032' }
    });

    const staffUser = await User.create({
      name: 'Ravi Shankar (Staff)',
      email: 'staff@grocery.com',
      password: 'staff123',
      role: 'staff',
      phone: '+91 9840654321',
      address: { street: '45 Velachery Main Rd', city: 'Chennai', state: 'Tamil Nadu', zipCode: '600042' }
    });

    const deliveryUser = await User.create({
      name: 'Karthik Raja (Delivery Partner)',
      email: 'delivery@grocery.com',
      password: 'delivery123',
      role: 'delivery',
      phone: '+91 9840998877',
      address: { street: '88 OMR Highway', city: 'Chennai', state: 'Tamil Nadu', zipCode: '600096' }
    });

    const customerUser = await User.create({
      name: 'Priya Sundaram',
      email: 'customer@gmail.com',
      password: 'customer123',
      role: 'customer',
      phone: '+91 9840112233',
      address: { street: 'Flat 4B, Green Acres, Adyar', city: 'Chennai', state: 'Tamil Nadu', zipCode: '600020' }
    });

    console.log('[Seed] Creating Suppliers...');
    const suppliers = await Supplier.insertMany([
      {
        name: 'Apex Dairy & Poultry Supplies',
        contactPerson: 'M. Anand',
        email: 'orders@apexdairy.in',
        phone: '+91 9444101010',
        address: 'Madhavaram Milk Colony, Chennai',
        categoriesSupplied: ['Dairy & Eggs'],
        rating: 4.8,
        averageDeliveryDays: 1,
        reliabilityScore: 98
      },
      {
        name: 'Golden Harvest Agro Mills',
        contactPerson: 'S. Narayanan',
        email: 'sales@goldenharvest.com',
        phone: '+91 9444202020',
        address: 'Koyambedu Wholesale Complex, Chennai',
        categoriesSupplied: ['Grains & Staples'],
        rating: 4.6,
        averageDeliveryDays: 2,
        reliabilityScore: 94
      },
      {
        name: 'Sunrise Bakeries & Snacks Co.',
        contactPerson: 'Deepa Krishnan',
        email: 'supply@sunrisebakers.com',
        phone: '+91 9444303030',
        address: 'Ambattur Industrial Estate, Chennai',
        categoriesSupplied: ['Bakery & Bread', 'Snacks & Packaged'],
        rating: 4.7,
        averageDeliveryDays: 1,
        reliabilityScore: 96
      },
      {
        name: 'Nilgiri Fresh Organics',
        contactPerson: 'K. Balaji',
        email: 'produce@nilgiriorganics.in',
        phone: '+91 9444404040',
        address: 'Fruit & Vegetable Market, Koyambedu, Chennai',
        categoriesSupplied: ['Fresh Produce'],
        rating: 4.9,
        averageDeliveryDays: 1,
        reliabilityScore: 99
      },
      {
        name: 'Southern FMCG & Beverage Distributors',
        contactPerson: 'V. Raman',
        email: 'sales@southernfmcg.com',
        phone: '+91 9444505050',
        address: 'Parrys Corner, Chennai',
        categoriesSupplied: ['Beverages', 'Personal & Home Care'],
        rating: 4.5,
        averageDeliveryDays: 3,
        reliabilityScore: 92
      }
    ]);

    const supplierMap = {
      'Dairy & Eggs': suppliers[0]._id,
      'Grains & Staples': suppliers[1]._id,
      'Bakery & Bread': suppliers[2]._id,
      'Snacks & Packaged': suppliers[2]._id,
      'Fresh Produce': suppliers[3]._id,
      'Beverages': suppliers[4]._id,
      'Personal & Home Care': suppliers[4]._id
    };

    console.log('[Seed] Creating Products...');
    const rawProducts = [
      {
        productId: 'PRD-1001',
        name: 'Fresh Pasteurised Whole Milk (1L)',
        category: 'Dairy & Eggs',
        brand: 'Aavin / Amul',
        description: 'Fresh toned rich farm milk pasteurized and packaged daily.',
        purchasePrice: 42,
        sellingPrice: 56,
        currentStock: 18, // near minStock
        minStockLevel: 20,
        maxStockLevel: 150,
        unit: 'packet',
        barcode: '8901262010015',
        imageUrl: 'https://images.unsplash.com/photo-1550583724-b2692b85b150?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1002',
        name: 'Whole Wheat Sandwich Bread (400g)',
        category: 'Bakery & Bread',
        brand: 'Modern / Britania',
        description: '100% whole grain brown bread, soft & fibre-rich.',
        purchasePrice: 32,
        sellingPrice: 48,
        currentStock: 12, // Critical low stock
        minStockLevel: 15,
        maxStockLevel: 80,
        unit: 'packet',
        barcode: '8901262010022',
        imageUrl: 'https://images.unsplash.com/photo-1509440159596-0249088772ff?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1003',
        name: 'Amul Pasteurised Butter (500g)',
        category: 'Dairy & Eggs',
        brand: 'Amul',
        description: 'Delicious utterly butterly Amul salted pure butter block.',
        purchasePrice: 220,
        sellingPrice: 275,
        currentStock: 35,
        minStockLevel: 10,
        maxStockLevel: 100,
        unit: 'pack',
        barcode: '8901262010039',
        imageUrl: 'https://images.unsplash.com/photo-1589985270826-4b7bb135bc9d?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1004',
        name: 'Farm Fresh Brown Eggs (Pack of 12)',
        category: 'Dairy & Eggs',
        brand: 'FreshFarm',
        description: 'Cage-free nutritious brown eggs rich in Omega 3 and protein.',
        purchasePrice: 75,
        sellingPrice: 98,
        currentStock: 45,
        minStockLevel: 15,
        maxStockLevel: 120,
        unit: 'tray',
        barcode: '8901262010046',
        imageUrl: 'https://images.unsplash.com/photo-1582722872445-44dc5f7e3c8f?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1005',
        name: 'Royal Kohinoor Basmati Rice (5kg)',
        category: 'Grains & Staples',
        brand: 'Kohinoor',
        description: 'Extra long grain aged aromatic Basmati rice for premium biryani & pulao.',
        purchasePrice: 450,
        sellingPrice: 580,
        currentStock: 50,
        minStockLevel: 10,
        maxStockLevel: 150,
        unit: 'bag',
        barcode: '8901262010053',
        imageUrl: 'https://images.unsplash.com/photo-1586201375761-83865001e31c?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1006',
        name: 'Organic Unpolished Toor Dal (1kg)',
        category: 'Grains & Staples',
        brand: 'Tata Sampann',
        description: 'Protein-rich natural unpolished yellow pigeon peas.',
        purchasePrice: 135,
        sellingPrice: 175,
        currentStock: 60,
        minStockLevel: 15,
        maxStockLevel: 150,
        unit: 'kg',
        barcode: '8901262010060',
        imageUrl: 'https://images.unsplash.com/photo-1514733670139-4d87a1941d55?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1007',
        name: 'Aashirvaad Shudh Chakki Atta (5kg)',
        category: 'Grains & Staples',
        brand: 'Aashirvaad',
        description: '100% whole wheat whole meal flour ground with traditional stone chakki.',
        purchasePrice: 215,
        sellingPrice: 260,
        currentStock: 40,
        minStockLevel: 12,
        maxStockLevel: 100,
        unit: 'bag',
        barcode: '8901262010077',
        imageUrl: 'https://images.unsplash.com/photo-1509440159596-0249088772ff?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1008',
        name: 'Tata Tea Gold Leaf (500g)',
        category: 'Beverages',
        brand: 'Tata Tea',
        description: 'Rich blend of aromatic CTC tea with gently rolled long leaves.',
        purchasePrice: 240,
        sellingPrice: 310,
        currentStock: 30,
        minStockLevel: 8,
        maxStockLevel: 80,
        unit: 'pack',
        barcode: '8901262010084',
        imageUrl: 'https://images.unsplash.com/photo-1576092768241-dec231879fc3?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1009',
        name: 'Nescafe Classic Instant Coffee (100g)',
        category: 'Beverages',
        brand: 'Nescafe',
        description: 'Signature rich roast blend of pure arabica & robusta coffee.',
        purchasePrice: 210,
        sellingPrice: 265,
        currentStock: 25,
        minStockLevel: 8,
        maxStockLevel: 60,
        unit: 'jar',
        barcode: '8901262010091',
        imageUrl: 'https://images.unsplash.com/photo-1514432324607-a09d9b4aefdd?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1010',
        name: 'Crisp Royal Gala Apples (1kg)',
        category: 'Fresh Produce',
        brand: 'Kashmir Orchard',
        description: 'Sweet, crisp and juicy handpicked mountain apples.',
        purchasePrice: 140,
        sellingPrice: 195,
        currentStock: 30,
        minStockLevel: 15,
        maxStockLevel: 90,
        unit: 'kg',
        barcode: '8901262010107',
        imageUrl: 'https://images.unsplash.com/photo-1560806887-1e4cd0b6cbd6?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1011',
        name: 'Robusta Golden Bananas (1 Dozen)',
        category: 'Fresh Produce',
        brand: 'Trichy Farms',
        description: 'Naturally ripened, potassium packed golden bananas.',
        purchasePrice: 35,
        sellingPrice: 55,
        currentStock: 15, // near expiry item
        minStockLevel: 10,
        maxStockLevel: 60,
        unit: 'dozen',
        barcode: '8901262010114',
        imageUrl: 'https://images.unsplash.com/photo-1571771894821-ce9b6c11b08e?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1012',
        name: 'Farm Fresh Vine Tomatoes (1kg)',
        category: 'Fresh Produce',
        brand: 'Dharmapuri Organics',
        description: 'Red, plump and firm garden tomatoes ideal for curries and salads.',
        purchasePrice: 28,
        sellingPrice: 42,
        currentStock: 50,
        minStockLevel: 20,
        maxStockLevel: 120,
        unit: 'kg',
        barcode: '8901262010121',
        imageUrl: 'https://images.unsplash.com/photo-1592924357228-91a4daadcfea?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1013',
        name: 'Barilla Italian Penne Rigate (500g)',
        category: 'Snacks & Packaged',
        brand: 'Barilla',
        description: 'Traditional 100% durum wheat semolina Italian pasta.',
        purchasePrice: 175,
        sellingPrice: 235,
        currentStock: 28,
        minStockLevel: 10,
        maxStockLevel: 60,
        unit: 'box',
        barcode: '8901262010138',
        imageUrl: 'https://images.unsplash.com/photo-1621996346565-e3d5d6281691?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1014',
        name: 'Del Monte Herb Pasta Sauce (500g)',
        category: 'Snacks & Packaged',
        brand: 'Del Monte',
        description: 'Authentic tomato sauce simmered with oregano, basil and extra virgin olive oil.',
        purchasePrice: 120,
        sellingPrice: 165,
        currentStock: 22,
        minStockLevel: 8,
        maxStockLevel: 50,
        unit: 'jar',
        barcode: '8901262010145',
        imageUrl: 'https://images.unsplash.com/photo-1572441713132-c542fc4fe282?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1015',
        name: 'Fortune Sunlite Refined Sunflower Oil (1L)',
        category: 'Grains & Staples',
        brand: 'Fortune',
        description: 'Light, healthy and fortified with vitamins A, D & E.',
        purchasePrice: 118,
        sellingPrice: 145,
        currentStock: 40,
        minStockLevel: 15,
        maxStockLevel: 100,
        unit: 'pouch',
        barcode: '8901262010152',
        imageUrl: 'https://images.unsplash.com/photo-1474979266404-7eaacbcd87c5?w=500&auto=format&fit=crop&q=60'
      },
      {
        productId: 'PRD-1016',
        name: 'Dettol Germ Protection Handwash (250ml)',
        category: 'Personal & Home Care',
        brand: 'Dettol',
        description: 'Trusted 99.9% antibacterial liquid hand wash.',
        purchasePrice: 70,
        sellingPrice: 95,
        currentStock: 35,
        minStockLevel: 10,
        maxStockLevel: 80,
        unit: 'pump bottle',
        barcode: '8901262010169',
        imageUrl: 'https://images.unsplash.com/photo-1584308666744-24d5c474f2ae?w=500&auto=format&fit=crop&q=60'
      }
    ];

    const insertedProducts = await Product.insertMany(rawProducts);
    console.log(`[Seed] Inserted ${insertedProducts.length} products.`);

    console.log('[Seed] Generating FEFO Batches with varying shelf lives...');
    const now = new Date();
    const batchesToInsert = [];

    for (const prod of insertedProducts) {
      const suppId = supplierMap[prod.category] || suppliers[0]._id;

      // Special batch rules:
      // Bananas and Bread have near-expiry batches (< 7 days) to demonstrate Section 6 alerts and Section 12 dynamic discounts!
      if (prod.productId === 'PRD-1011') { // Bananas
        const exp = new Date(now.getTime() + 4 * 24 * 60 * 60 * 1000); // 4 days remaining!
        batchesToInsert.push({
          batchNumber: 'BATCH-BAN-CRIT-01',
          product: prod._id,
          productIdStr: prod.productId,
          manufacturingDate: new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000),
          expiryDate: exp,
          initialQuantity: 15,
          currentQuantity: 15,
          costPrice: prod.purchasePrice,
          supplier: suppId,
          locationRack: 'PRODUCE-RACK-01',
          status: 'CRITICAL_EXPIRY_7D'
        });
      } else if (prod.productId === 'PRD-1002') { // Bread
        const exp = new Date(now.getTime() + 5 * 24 * 60 * 60 * 1000); // 5 days remaining!
        batchesToInsert.push({
          batchNumber: 'BATCH-BRD-CRIT-02',
          product: prod._id,
          productIdStr: prod.productId,
          manufacturingDate: new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000),
          expiryDate: exp,
          initialQuantity: 12,
          currentQuantity: 12,
          costPrice: prod.purchasePrice,
          supplier: suppId,
          locationRack: 'BAKERY-RACK-02',
          status: 'CRITICAL_EXPIRY_7D'
        });
      } else if (prod.productId === 'PRD-1001') { // Milk
        // Two batches to demonstrate FEFO ordering!
        const expEarly = new Date(now.getTime() + 6 * 24 * 60 * 60 * 1000); // Earliest batch
        const expLate = new Date(now.getTime() + 18 * 24 * 60 * 60 * 1000);
        batchesToInsert.push({
          batchNumber: 'BATCH-MLK-EARLY-01',
          product: prod._id,
          productIdStr: prod.productId,
          manufacturingDate: new Date(now.getTime() - 1 * 24 * 60 * 60 * 1000),
          expiryDate: expEarly,
          initialQuantity: 10,
          currentQuantity: 10,
          costPrice: prod.purchasePrice,
          supplier: suppId,
          locationRack: 'CHILLER-01',
          status: 'CRITICAL_EXPIRY_7D'
        });
        batchesToInsert.push({
          batchNumber: 'BATCH-MLK-LATER-02',
          product: prod._id,
          productIdStr: prod.productId,
          manufacturingDate: now,
          expiryDate: expLate,
          initialQuantity: 8,
          currentQuantity: 8,
          costPrice: prod.purchasePrice,
          supplier: suppId,
          locationRack: 'CHILLER-01',
          status: 'NEAR_EXPIRY_30D'
        });
      } else {
        // Standard shelf life batch (60 to 180 days)
        const expDays = Math.floor(60 + Math.random() * 120);
        const exp = new Date(now.getTime() + expDays * 24 * 60 * 60 * 1000);
        batchesToInsert.push({
          batchNumber: `BATCH-${prod.productId.replace('PRD-', '')}-${Date.now().toString().slice(-4)}`,
          product: prod._id,
          productIdStr: prod.productId,
          manufacturingDate: new Date(now.getTime() - 10 * 24 * 60 * 60 * 1000),
          expiryDate: exp,
          initialQuantity: prod.currentStock,
          currentQuantity: prod.currentStock,
          costPrice: prod.purchasePrice,
          supplier: suppId,
          locationRack: `AISLE-${prod.category[0]}-0${Math.floor(1 + Math.random() * 4)}`,
          status: 'ACTIVE'
        });
      }
    }

    await Batch.insertMany(batchesToInsert);
    console.log(`[Seed] Created ${batchesToInsert.length} batches.`);

    console.log('[Seed] Generating 90-day historical daily sales data for AI Model training...');
    const salesRecords = [];
    for (let dayOffset = 90; dayOffset >= 0; dayOffset--) {
      const d = new Date(now.getTime() - dayOffset * 24 * 60 * 60 * 1000);
      const dateStr = d.toISOString().split('T')[0];
      const dow = d.getDay();
      const month = d.getMonth() + 1;
      const isWeekend = (dow === 0 || dow === 6);

      for (const prod of insertedProducts) {
        // Base sales quantity
        let baseQty = 8;
        if (prod.category === 'Dairy & Eggs' || prod.category === 'Bakery & Bread') baseQty = 14;
        if (prod.category === 'Fresh Produce') baseQty = 12;

        const weekendMultiplier = isWeekend ? 1.35 : 0.9;
        const randomFactor = 0.8 + Math.random() * 0.4;
        const unitsSold = Math.max(1, Math.round(baseQty * weekendMultiplier * randomFactor));
        const revenue = unitsSold * prod.sellingPrice;
        const cost = unitsSold * prod.purchasePrice;
        const profit = revenue - cost;

        salesRecords.push({
          product: prod._id,
          productIdStr: prod.productId,
          date: dateStr,
          unitsSold,
          revenue,
          cost,
          profit,
          dayOfWeek: dow,
          month
        });
      }
    }

    await DailySale.insertMany(salesRecords);
    console.log(`[Seed] Seeded ${salesRecords.length} historical daily sale data points.`);

    console.log('[Seed] Generating initial Purchase Orders and Auto-Reorder recommendation...');
    // PRD-1002 (Bread) is below min stock level (12 <= 15)
    const breadProd = insertedProducts.find(p => p.productId === 'PRD-1002');
    if (breadProd) {
      const poQty = 50;
      const poTotal = poQty * breadProd.purchasePrice;
      await PurchaseOrder.create({
        poNumber: 'PO-REC-10029',
        supplier: suppliers[2]._id,
        supplierName: suppliers[2].name,
        items: [{
          product: breadProd._id,
          productIdStr: breadProd.productId,
          productName: breadProd.name,
          recommendedQuantity: poQty,
          orderQuantity: poQty,
          estimatedUnitCost: breadProd.purchasePrice,
          totalCost: poTotal
        }],
        totalCost: poTotal,
        status: 'RECOMMENDED_BY_AI',
        triggerReason: `Stock dropped to ${breadProd.currentStock} (below min stock ${breadProd.minStockLevel})`,
        expectedDeliveryDate: new Date(Date.now() + 24 * 60 * 60 * 1000)
      });
    }

    console.log('[Seed] Generating Sample Customer Orders...');
    const milk = insertedProducts.find(p => p.productId === 'PRD-1001');
    const bread = insertedProducts.find(p => p.productId === 'PRD-1002');
    const butter = insertedProducts.find(p => p.productId === 'PRD-1003');

    // Sample Delivered Order
    await Order.create({
      orderNumber: 'ORD-7001',
      customer: customerUser._id,
      customerDetails: {
        name: customerUser.name,
        phone: customerUser.phone,
        address: customerUser.address.street + ', Chennai'
      },
      items: [
        {
          product: milk._id,
          productIdStr: milk.productId,
          productName: milk.name,
          quantity: 2,
          unitPrice: milk.sellingPrice,
          purchasePrice: milk.purchasePrice,
          subtotal: milk.sellingPrice * 2
        },
        {
          product: bread._id,
          productIdStr: bread.productId,
          productName: bread.name,
          quantity: 1,
          unitPrice: bread.sellingPrice,
          purchasePrice: bread.purchasePrice,
          subtotal: bread.sellingPrice
        }
      ],
      subtotal: 160,
      discountAmount: 0,
      taxAmount: 8,
      deliveryFee: 40,
      finalTotal: 208,
      profitEarned: 52,
      paymentMethod: 'UPI',
      paymentStatus: 'COMPLETED',
      transactionId: 'TXN-UPI-994827',
      orderStatus: 'Delivered',
      orderType: 'ONLINE',
      deliveryPartner: deliveryUser._id,
      deliveryPartnerName: deliveryUser.name,
      deliveredAt: new Date(Date.now() - 2 * 60 * 60 * 1000),
      statusTimeline: [
        { status: 'Placed', timestamp: new Date(Date.now() - 4 * 60 * 60 * 1000), notes: 'Order placed by customer' },
        { status: 'Confirmed', timestamp: new Date(Date.now() - 3.5 * 60 * 60 * 1000), notes: 'Order confirmed by store' },
        { status: 'Packed', timestamp: new Date(Date.now() - 3 * 60 * 60 * 1000), notes: 'FEFO stock packed' },
        { status: 'Out for Delivery', timestamp: new Date(Date.now() - 2.5 * 60 * 60 * 1000), notes: 'Dispatched with Karthik' },
        { status: 'Delivered', timestamp: new Date(Date.now() - 2 * 60 * 60 * 1000), notes: 'Delivered safely to doorstep' }
      ]
    });

    // Sample Pending Order (Section 8 & 10)
    await Order.create({
      orderNumber: 'ORD-7002',
      customer: customerUser._id,
      customerDetails: {
        name: customerUser.name,
        phone: customerUser.phone,
        address: customerUser.address.street + ', Chennai'
      },
      items: [
        {
          product: butter._id,
          productIdStr: butter.productId,
          productName: butter.name,
          quantity: 1,
          unitPrice: butter.sellingPrice,
          purchasePrice: butter.purchasePrice,
          subtotal: butter.sellingPrice
        },
        {
          product: bread._id,
          productIdStr: bread.productId,
          productName: bread.name,
          quantity: 1,
          unitPrice: bread.sellingPrice,
          purchasePrice: bread.purchasePrice,
          subtotal: bread.sellingPrice
        }
      ],
      subtotal: 323,
      discountAmount: 20,
      taxAmount: 16,
      deliveryFee: 40,
      finalTotal: 359,
      profitEarned: 71,
      paymentMethod: 'CARD',
      paymentStatus: 'COMPLETED',
      transactionId: 'TXN-CARD-881234',
      orderStatus: 'Placed',
      orderType: 'ONLINE',
      statusTimeline: [
        { status: 'Placed', timestamp: new Date(), notes: 'Customer paid via Credit Card' }
      ]
    });

    // Audit logs for initial setup
    await AuditLog.create({
      user: adminUser._id,
      userName: adminUser.name,
      userEmail: adminUser.email,
      userRole: adminUser.role,
      action: 'LOGIN',
      details: 'System initialized and seed data successfully loaded'
    });

    console.log('✅ [Seed] Database seeding completed successfully!');
    return true;
  } catch (err) {
    console.error('❌ [Seed Error]', err);
    throw err;
  }
};

module.exports = { seedDatabase };

if (require.main === module) {
  seedDatabase()
    .then(() => process.exit(0))
    .catch(() => process.exit(1));
}
