# THREDrDB - TypeScript Excel Add-in for Real-time Data Processing

## Overview
THREDrDB is a modern Excel add-in that provides powerful custom functions for real-time data processing, API integration, and database operations. Originally converted from VB.NET ExcelDNA, it now leverages TypeScript and Office.js for enhanced performance and streaming capabilities.

## 🚀 Key Features

### Real-time Data Streaming
- **Live Updates**: Streaming custom functions with automatic cleanup
- **API Monitoring**: Real-time API endpoint monitoring
- **Clock & Timers**: Precise timing functions for triggers and displays
- **Data Simulation**: Generate sample data for testing and development

### API Integration
- **Multiple Authentication**: Bearer, Basic, API Key, and custom auth methods
- **JSON Processing**: Advanced JSON path extraction and transformation
- **Caching System**: Intelligent caching to prevent excessive API calls
- **URL Building**: Dynamic URL construction with proper encoding

### Database Operations
- **Local SQLite**: Client-side database with workbook isolation
- **Full SQL Support**: Execute complex queries with aggregation and joins
- **Data Loading**: Load Excel ranges directly into database tables
- **Pagination**: Efficient handling of large datasets

### Utility Functions
- **Data Transformation**: Convert JSON to various Excel-compatible formats
- **Error Handling**: Comprehensive error handling with debugging modes
- **Performance Optimization**: Automatic resource management and cleanup

## 🔧 Installation

### Prerequisites
- Microsoft Excel (Desktop or Online)
- Node.js 16+ (for development)
- TypeScript 4+ (for development)

### Quick Start
1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/THREDrDB.git
   cd THREDrDB
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Build the project**
   ```bash
   npm run build:dev
   ```

4. **Start development server**
   ```bash
   npm run dev-server
   ```

5. **Load in Excel**
   - Open Excel
   - Go to Insert > Office Add-ins > Upload My Add-in
   - Select the `manifest.xml` file

### Test Installation
In Excel, try this function to verify everything is working:
```excel
=TEST()
```
**Expected result:** "Custom functions are working!"

## 📊 Quick Examples

### Real-time Clock
```excel
=CLOCK(1)  // Updates every second
```

### API Data Retrieval
```excel
=APICALL("https://jsonplaceholder.typicode.com/users")
```

### Database Operations
```excel
// Load Excel data to database
=REFRESHTABLE(A1:C10, "my_table")

// Query the database
=QUERY("SELECT * FROM my_table WHERE column1 > 100", TRUE)
```

### Real-time API Monitoring
```excel
=REALTIMEAPI("https://api.example.com/data", 30, "results", "", "live", "bearer", "your-token")
```

## 📚 Documentation

### Core Documentation
- **[API Reference](API_REFERENCE.md)** - Complete function documentation
- **[Examples & Tutorials](EXAMPLES.md)** - Real-world usage examples
- **[Development Notes](DEVELOPMENT_NOTES.md)** - Quick reference guide
- **[Chat History](CHAT_HISTORY.md)** - Development conversation record
- **[Changelog](CHANGELOG.md)** - Version history and updates

### Function Categories
| Category | Functions | Description |
|----------|-----------|-------------|
| **Streaming** | CLOCK, TIMER, REALTIMEAPI, DATASIMULATOR | Real-time updates with automatic cleanup |
| **API** | APICALL, APICALLAUTH, APICALLENHANCED, URLBUILDER | HTTP requests with authentication |
| **Database** | REFRESHTABLE, QUERY, VOLATILEQUERY, SUMFIELD, LOOKUP, DATAPAGES | SQLite operations |
| **Utility** | TRANSFORMJSON, TEST | Data transformation and testing |

## 🏗️ Architecture

### Technology Stack
- **TypeScript**: Type-safe JavaScript development
- **Office.js**: Microsoft Office JavaScript API
- **SQLite.js**: Client-side database operations
- **Webpack**: Module bundling and build process

### Key Design Principles
- **Streaming First**: Real-time updates using CustomFunctions.StreamingInvocation
- **Resource Management**: Automatic cleanup and memory management
- **Error Handling**: Comprehensive error handling with debugging modes
- **Performance**: Caching, pagination, and efficient data processing

### Project Structure
```
THREDrDB/
├── src/
│   ├── functions/
│   │   └── functions.ts          # Core custom functions
│   ├── commands/
│   │   └── commands.ts           # Ribbon commands
│   ├── taskpane/
│   │   ├── taskpane.html         # Task pane UI
│   │   ├── taskpane.ts           # Task pane logic
│   │   └── taskpane.css          # Task pane styles
│   └── database/                 # Database utilities
├── assets/                       # Icons and images
├── scripts/                      # Build and deployment scripts
├── manifest.xml                  # Add-in manifest
├── package.json                  # Dependencies and scripts
├── webpack.config.js             # Build configuration
└── tsconfig.json                 # TypeScript configuration
```

## 🛠️ Development

### Available Scripts
```bash
npm run build:dev      # Development build
npm run build          # Production build
npm run dev-server     # Start development server
npm run start:desktop  # Debug in Excel Desktop
npm run lint           # Check code quality
npm run lint:fix       # Fix linting issues
```

### Development Workflow
1. **Start development server**: `npm run dev-server`
2. **Make changes** to source files
3. **Test in Excel** using the development server
4. **Build for production**: `npm run build`

### Debugging
- Use `debugging=TRUE` in function parameters
- Check browser console for detailed logs
- Use `=TEST()` to verify function loading
- Enable debugging mode during development

## 🔐 Security

### Authentication
- Support for multiple authentication methods
- Secure token handling (not stored permanently)
- HTTPS recommended for all API endpoints

### Data Protection
- Local SQLite database (no network access)
- Workbook-specific data isolation
- Proper input validation and sanitization

## 🚀 Performance

### Optimization Features
- **API Caching**: Reduces redundant requests (default 10 minutes)
- **Pagination**: Efficient handling of large datasets
- **Streaming**: Real-time updates without blocking
- **Resource Cleanup**: Automatic timer and memory management

### Best Practices
- Use appropriate cache timeouts for API calls
- Implement pagination for large datasets
- Use aggregation queries instead of loading all data
- Monitor API rate limits and costs

## 🤝 Contributing

### Development Setup
1. Fork the repository
2. Create a feature branch: `git checkout -b feature/amazing-feature`
3. Make your changes
4. Add tests if applicable
5. Commit your changes: `git commit -m 'Add amazing feature'`
6. Push to the branch: `git push origin feature/amazing-feature`
7. Open a Pull Request

### Code Style
- Follow TypeScript best practices
- Use ESLint configuration provided
- Add JSDoc comments for all functions
- Include error handling and validation

### Testing
- Test all functions in Excel Desktop and Online
- Verify streaming functions clean up properly
- Test error conditions and edge cases
- Document any new features or changes

## 🐛 Troubleshooting

### Common Issues
- **Function not found**: Restart Excel or reload add-in
- **Database errors**: Check table names and SQL syntax
- **API failures**: Verify URLs and authentication
- **Performance issues**: Enable caching and use pagination

### Getting Help
- Check the [Examples](EXAMPLES.md) for common usage patterns
- Review [API Reference](API_REFERENCE.md) for function details
- Enable debugging mode for detailed error messages
- Check browser console for additional information

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Microsoft Office.js team for streaming functions API
- SQLite.js project for client-side database support
- TypeScript and webpack communities for excellent tooling
- Contributors and beta testers for feedback and suggestions

## 📞 Support

For questions, issues, or feature requests:
- Open an issue on GitHub
- Check existing documentation
- Review examples and tutorials
- Enable debugging mode for detailed error information

---

**THREDrDB** - Bringing real-time data processing to Excel with modern web technologies.

*Last updated: July 15, 2025*
