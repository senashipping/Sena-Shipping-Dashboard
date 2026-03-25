require("dotenv").config();
const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const morgan = require("morgan");
const helmet = require("helmet");

// Import routes
const authRoutes = require("./routes/auth");
const adminRoutes = require("./routes/admin");
const formRoutes = require("./routes/forms");
const submissionRoutes = require("./routes/submissions");
const notificationRoutes = require("./routes/notifications");

// Import middleware
const { errorHandler, notFound } = require("./middleware/errorHandler");

// Import services
const { initializeScheduledTasks } = require("./services/cronService");
const { seedDefaultData } = require("./utils/seedData");

const app = express();

// Security middleware
app.use(helmet());

// Logging middleware
app.use(morgan("dev"));

// CORS configuration - Allow all origins for development
app.use(
  cors({
    origin: "*",
    credentials: false, // Set to false when using wildcard origin
    methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization", "X-Requested-With"],
  })
);

// Body parsing middleware
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true, limit: "50mb" }));

// API Routes
app.use("/api/auth", authRoutes);
app.use("/api/admin", adminRoutes);
app.use("/api/forms", formRoutes);
app.use("/api/submissions", submissionRoutes);
app.use("/api/notifications", notificationRoutes);

// Health check endpoint
app.get("/api/health", (req, res) => {
  res.json({
    success: true,
    message: "Ship Dashboard API is running",
    timestamp: new Date().toISOString(),
    environment: process.env.NODE_ENV || "development",
  });
});

// Root endpoint
app.get("/", (req, res) => {
  res.json({
    success: true,
    message: "Welcome to Ship Dashboard API",
    version: "1.0.0",
    documentation: "/api/docs",
  });
});

// Error handling middleware
app.use(notFound);
app.use(errorHandler);

const PORT = process.env.PORT || 8080;

mongoose
  .connect(process.env.MONGO_URI)
  .then(async () => {
    console.log("✓ MongoDB connected successfully");

    // Seed default data in development
    if (process.env.NODE_ENV !== "production") {
      await seedDefaultData();
    }

    // Initialize scheduled tasks
    initializeScheduledTasks();

    app.listen(PORT, () => {
      console.log(`✓ Server running on port ${PORT}`);
      console.log(`✓ Environment: ${process.env.NODE_ENV || "development"}`);
      console.log(`✓ API Health Check: http://localhost:${PORT}/api/health`);
    });
  })
  .catch((err) => {
    console.error("❌ MongoDB connection error:", err);
    process.exit(1);
  });
