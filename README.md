# 💰 Personal Loan Tracker - Pakistan Edition

A comprehensive, professional-grade loan tracking system built for Google Sheets with unlimited payment tracking, live dashboard analytics, and full Pakistani currency support.

## 🌟 Features

### ✨ Core Functionality
- **Unlimited Payments**: Track as many installments as needed per loan
- **Live Dashboard**: Real-time analytics and statistics in Google Sheets
- **Payment History**: Complete transaction log with unique IDs
- **Pakistani Currency**: Native support for lac, crore, k formats
- **Modern Forms**: Professional HTML forms with validation

### 📊 Dashboard Analytics
- Total borrowed vs paid vs remaining amounts
- Loan purpose breakdown with percentages
- Active loans summary with progress indicators
- Recent payments tracking (last 10 transactions)
- Payment frequency and pattern analysis

### 🎯 Advanced Features
- Unique loan and payment IDs for tracking
- Payment method recording (Cash, Bank Transfer, Mobile Banking)
- Automatic status management (Active → Paid)
- Color-coded visual indicators
- Auto-refresh calculations

## 🚀 Quick Start

### Prerequisites
- Google Account
- Google Sheets access
- Basic familiarity with Google Apps Script

### Installation
1. **Create New Spreadsheet**
Go to sheets.google.com → Create new spreadsheet

2. **Open Apps Script**
Extensions → Apps Script

3. **Add the Code**
Copy the code from src/Code.gs
Paste into the Apps Script editor
Save the project

4. **Initialize Tracker**
Refresh your spreadsheet
Use menu: 💰 Enhanced Loan Tracker → 🏗️ Initialize Loan Tracker

## 📖 Usage Guide

### Adding Loans
1. Go to **💰 Enhanced Loan Tracker** menu
2. Click **➕ Add New Loan**
3. Fill in the modern form:
- Lender name
- Amount (supports: 5 lac, 50k, 2.5 lac, 150000)
- Purpose (optional)
- Expected monthly payment (optional)
- Notes (optional)

### Recording Payments
1. Click **💸 Record Payment**
2. Select the loan from dropdown
3. Enter payment details:
- Amount (flexible format)
- Date (auto-filled with today)
- Payment method
- Notes (optional)

### Viewing Analytics
- **Dashboard Sheet**: Live statistics and analytics
- **Payment History**: Complete transaction log
- **Refresh Dashboard**: Manual refresh option available

## 💸 Amount Format Examples

| Input | Output |
|-------|--------|
| `5 lac` | ₨5,00,000 |
| `2.5 lac` | ₨2,50,000 |
| `50k` | ₨50,000 |
| `1.2 cr` | ₨1,20,00,000 |
| `150000` | ₨1,50,000 |

## 📊 Dashboard Features

### Key Statistics Section
- Total loans count
- Active vs paid loans
- Total borrowed amount
- Total paid amount
- Remaining balance
- Payment count and averages
- Progress percentage

### Analytics Sections
- **Loan Purpose Breakdown**: Visual breakdown by category
- **Active Loans Summary**: Current loan status with color coding
- **Recent Payments**: Last 10 transactions with trends

## 🛠️ Technical Details

### Architecture
- **3-Sheet System**:
- `Loan Tracker`: Main loan summary
- `Payment History`: Complete payment log
- `Dashboard`: Live analytics
- **Google Apps Script**: Server-side logic
- **HTML Forms**: Modern UI components
- **Automatic Calculations**: Real-time updates

### Data Structure
- **Unique IDs**: Each loan and payment gets unique identifier
- **Relational Links**: Payments linked to loans via IDs
- **Status Tracking**: Automatic status updates
- **Audit Trail**: Complete history preservation

## 🎯 Use Cases

### Personal Finance
- Track personal loans from friends/family
- Monitor repayment progress
- Analyze payment patterns

### Small Business
- Manage business loans
- Track multiple lenders
- Generate payment reports

### Family Management
- Family loan coordination
- Shared financial planning
- Transparent payment tracking

## 🌍 Localization

### Pakistani Market Features
- **Currency**: Pakistani Rupee (PKR) formatting
- **Numbers**: Local number system (lac, crore)
- **Dates**: DD/MM/YYYY format
- **Language**: English with local terminology


---

**Made with ❤️ for the Pakistani community**

*Simplifying personal finance management with modern tools*
