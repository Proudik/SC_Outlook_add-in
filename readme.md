# SingleCase Outlook Add-in

A Microsoft Outlook add-in that enables users to seamlessly file emails and attachments from Outlook directly into cases in SingleCase.

## 🎯 Features

### Core Functionality
- **Smart Case Filing**: File emails to SingleCase cases directly from Outlook (both read and compose modes)
- **Intelligent Suggestions**: AI-powered case suggestions based on:
  - Email content analysis
  - Historical filing patterns
  - Participant history
  - Conversation context
- **Attachment Management**: Selectively choose which attachments to include when filing
- **Internal Email Detection**: Automatically detects internal-only emails and provides guidance

### User Experience
- **Clear Selection UI**: Prominently displays selected cases with multiple visual indicators:
  - Checkmark icons
  - "SELECTED" badges
  - Blue gradient highlighting
  - Strong borders and shadows
- **Real-time Status**: Shows filing status with visual feedback
- **Already Filed Detection**: Recognizes previously filed emails and displays their case association
- **Document Management**: View, rename, and delete filed documents directly from Outlook

### Smart Features
- **Auto-file on Send**: Option to automatically file outgoing emails to selected cases
- **Frequent Case Detection**: Identifies commonly used cases for quick access
- **Reply Context**: Maintains case context when replying to emails
- **Version Management**: Upload new versions of existing documents
- **Duplicate Prevention**: Detects already filed emails to prevent duplicates

## 🚀 Getting Started

### Prerequisites
- Node.js (v16 or higher)
- Microsoft Outlook (Desktop, Web, or Mac)
- Access to a SingleCase instance

### Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd "SC_Outlook_addin v2 copy"
```

2. Install dependencies:
```bash
npm install
```

3. Generate development certificates:
```bash
npm run dev-server
```

4. Start the development server:
```bash
npm run dev-server
```

5. Sideload the add-in in Outlook:
```bash
npm start
```

## 🛠️ Development

### Available Scripts

- `npm run build` - Build for production
- `npm run build:dev` - Build for development
- `npm run dev-server` - Start development server with hot reload
- `npm run watch` - Watch mode for development
- `npm start` - Start debugging in Outlook
- `npm stop` - Stop debugging
- `npm run validate` - Validate manifest.xml
- `npm run lint` - Run ESLint
- `npm run lint:fix` - Fix ESLint issues
- `npm run prettier` - Format code with Prettier

### Project Structure

```
├── src/
│   ├── taskpane/
│   │   ├── components/
│   │   │   ├── MainWorkspace/      # Main filing interface
│   │   │   ├── CaseSelector/       # Case selection component
│   │   │   ├── PromptBubble/       # User prompts and messages
│   │   │   └── SettingsModal/      # User settings
│   │   └── taskpane.tsx            # Main entry point
│   ├── commands/                    # Command handlers
│   ├── services/                    # API services
│   ├── utils/                       # Utility functions
│   └── storage/                     # Local storage management
├── manifest.xml                     # Add-in manifest
├── webpack.config.js               # Webpack configuration
└── package.json                    # Dependencies and scripts
```

### Key Components

#### MainWorkspace
The main interface for filing emails. Handles:
- Email reading (received mode)
- Email composition (compose mode)
- Case selection and filing workflow
- Internal email detection and warnings
- Status messages and user feedback

#### CaseSelector
Advanced case selection UI with:
- Smart suggestions (content-based and history-based)
- Search functionality
- Favorite/all cases filtering
- Clear visual selection indicators

#### Case Suggestion Engine
Located in `src/utils/caseSuggestionEngine.ts`:
- Analyzes email content, participants, and context
- Provides confidence scores for suggestions
- Explains reasoning for each suggestion

## 🔐 Authentication

The add-in uses Microsoft Authentication Library (MSAL) for secure authentication:
- **SSO (Single Sign-On)**: Seamlessly authenticates using Office credentials
- **Token Management**: Automatic token refresh
- **Secure Storage**: Tokens stored securely in browser storage

## 🎨 UI/UX Improvements

### Recent Enhancements
- **Selected Case Section**: Prominent display of currently selected case at the top
- **Visual Indicators**: Multiple cues (checkmarks, badges, borders) for selection
- **No Duplication**: Selected cases filtered from suggestions to prevent confusion
- **Internal Email Guard**: Clear prompts for internal-only emails
- **Accessibility**: Form labels and ARIA attributes for screen readers

## 📋 Configuration

### Settings
Users can configure:
- Case list scope (Favorites or All)
- Auto-file on send preference
- Document filing options
- Attachment selection defaults

### Storage
The add-in uses:
- **Local Storage**: User preferences, filing history
- **Session Storage**: Temporary data during active sessions
- **Outlook Custom Properties**: Email filing metadata

## 🧪 Testing

### Manual Testing
1. Start the dev server: `npm run dev-server`
2. Launch Outlook: `npm start`
3. Test scenarios:
   - Filing received emails
   - Filing outgoing emails
   - Selecting different cases
   - Handling attachments
   - Internal email detection

### Debugging
- Use browser DevTools (F12 in Outlook desktop)
- Check console logs for detailed operation traces
- Manifest validation: `npm run validate`

## 📦 Building for Production

1. Update version in `manifest.xml`
2. Build the project:
```bash
npm run build
```

3. Deploy the dist folder to your hosting service
4. Update manifest URLs to production URLs
5. Submit to Microsoft AppSource (optional)

## 🔄 Recent Changes

### Version 1.0.0.3
- ✅ Enhanced case selection UI with clear visual indicators
- ✅ Fixed internal email detection in received mode
- ✅ Improved "File anyway" button behavior
- ✅ Added prominent "Selected Case" section
- ✅ Fixed badge overlap with long case names
- ✅ Accessibility improvements for form inputs

## 🐛 Known Issues

- None currently reported

## 📝 License

MIT

## 🤝 Contributing

1. Create a feature branch
2. Make your changes
3. Test thoroughly
4. Submit a pull request

## 📞 Support

For issues or questions, please contact the SingleCase support team.

---

**Built with React, TypeScript, and Office.js**
