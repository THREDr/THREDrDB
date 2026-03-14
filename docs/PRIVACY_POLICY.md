# Privacy Policy for THREDrDB

**Effective Date**: November 3, 2025  
**Last Updated**: November 3, 2025

---

## 1. Data Controller Information

**Organization Name**: Thredr LLC  
**Address**: [Your complete business address]  
**Email**: support@thredr.com  
**Website**: [Your website URL]  

For privacy-related inquiries or to exercise your data subject rights, please contact us at: support@thredr.com

---

## 2. Introduction

THREDrDB is a Microsoft Excel add-in that provides custom functions for database operations, data import, and SQL querying capabilities. This privacy policy explains how THREDrDB handles personal information in compliance with the General Data Protection Regulation (GDPR), the California Consumer Privacy Act (CCPA), and other applicable data protection laws.

We are committed to protecting your privacy and handling your data transparently and securely.

---

## 3. Scope and Application

This privacy policy applies to:
- THREDrDB Excel add-in (all versions)
- All users of THREDrDB worldwide
- All data processing activities conducted by or on behalf of Thredr LLC in connection with THREDrDB

---

## 4. Information We Collect

### 4.1 Personal Data

THREDrDB collects minimal personal data necessary to provide the add-in functionality:

**User Identification Data:**
- **Microsoft User Object ID (OID)**: A unique pseudonymous identifier assigned by Microsoft to your Microsoft 365 account. This identifier does not directly reveal your identity and is used solely for license verification and subscription management.

**We explicitly do NOT collect:**
- ❌ Email addresses
- ❌ Names
- ❌ Phone numbers
- ❌ Demographic information
- ❌ Location data
- ❌ Device identifiers
- ❌ IP addresses

### 4.2 Excel Workbook Data

THREDrDB processes data from your Excel workbooks to provide its core functionality:
- **Data Processed**: Excel cell data, formulas, table structures, and imported data that you choose to work with using THREDrDB functions
- **Processing Location**: All data processing occurs locally within your browser and Excel client
- **Storage**: No Excel workbook data is transmitted to external servers or stored persistently by THREDrDB
- **Retention**: Data is held in browser memory only during active use and is automatically cleared when you close Excel or the add-in

### 4.3 Authentication Data

- **Authentication Tokens**: Temporary tokens obtained from Microsoft Authentication Library (MSAL) for authentication purposes
- **Retention**: Tokens are session-based and are not stored persistently by THREDrDB
- **Purpose**: To verify your identity and enable secure access to the add-in

### 4.4 Technical Data

THREDrDB may collect limited technical information for functionality and troubleshooting:
- Add-in version number
- Error logs (if you choose to share them for support purposes)
- No automatic telemetry or usage analytics are collected

---

## 5. How We Use Your Information

### 5.1 Legal Basis for Processing

We process your personal data under the following legal bases:

**Legitimate Interest** (Article 6(1)(f) GDPR):
- Providing the core functionality of the THREDrDB add-in
- Verifying license and subscription status
- Improving and maintaining the add-in

**Consent** (Article 6(1)(a) GDPR):
- When you install and use the add-in, you consent to the processing described in this policy
- You may withdraw consent at any time by uninstalling the add-in

**Contractual Necessity** (Article 6(1)(b) GDPR):
- Processing necessary to provide the add-in service you have requested

### 5.2 Purpose of Data Processing

We use your User Object ID (OID) for:
- **License Verification**: Confirming your subscription status and entitlement to use THREDrDB features
- **Subscription Management**: Linking your Microsoft account to your subscription through our third-party payment provider (without revealing your identity to us)
- **Account Management**: Ensuring only authorized users can access premium features

We process your Excel workbook data for:
- **Core Functionality**: Executing SQL queries, importing data, and performing database operations as requested by you
- **Feature Delivery**: Providing the custom functions and features you use within Excel

---

## 6. Data Storage and Retention

### 6.1 Storage Location

- **User Object ID (OID)**: Not stored persistently by THREDrDB. Only used during active sessions for license verification with our subscription service provider.
- **Excel Workbook Data**: Processed entirely within your browser and Excel client. Never transmitted to external servers or stored by Thredr LLC.
- **Client-Side Database**: THREDrDB uses sql.js (SQLite compiled to JavaScript) which runs entirely in your browser memory. No data is uploaded to external servers.

### 6.2 Data Retention Periods

| Data Type | Retention Period | Rationale |
|-----------|------------------|-----------|
| User Object ID (OID) | Session only (not stored) | Used only for real-time license verification |
| Authentication Tokens | Session duration | Required for active authentication only |
| Excel Workbook Data | Browser memory only | Cleared automatically when add-in is closed |
| Support Communications | 3 years | For customer service and legal compliance |

### 6.3 Backup and Archival

Since THREDrDB does not persistently store user data, no backups of personal information are created. Our code repository backups contain no user data.

---

## 7. Data Sharing and Third-Party Services

### 7.1 Third Parties We Share Data With

We share your User Object ID (OID) with the following third parties:

**Subscription and Payment Service Provider** (To be determined):
- **Data Shared**: User Object ID (OID), subscription plan information
- **Purpose**: To manage subscriptions, process payments, and verify license entitlements
- **Legal Basis**: Contractual necessity
- **Safeguards**: Data Processing Agreement (DPA) with GDPR-compliant terms
- **Your Control**: Subscription data is managed by you through the payment provider's portal

**Microsoft Corporation**:
- **Data Shared**: Authentication requests (User.Read API permission)
- **Purpose**: To authenticate your identity and obtain your User Object ID
- **Legal Basis**: Contractual necessity (Microsoft 365 platform integration)
- **Safeguards**: Microsoft's own privacy policy and GDPR compliance

### 7.2 Third Parties We Do NOT Share Data With

We do not share your data with:
- ❌ Advertising networks
- ❌ Social media platforms
- ❌ Data brokers or analytics services
- ❌ Marketing companies
- ❌ Any other third parties not listed above

### 7.3 Data Processing Agreements

All third-party service providers with whom we share personal data are required to:
- Sign Data Processing Agreements (DPAs) with GDPR-compliant terms
- Implement appropriate technical and organizational security measures
- Process data only according to our instructions
- Notify us of any data breaches

---

## 8. International Data Transfers

### 8.1 Cross-Border Transfers

Your User Object ID may be processed by our subscription service provider, which may be located outside your country of residence. When data is transferred internationally, we ensure adequate protection through:

- **Standard Contractual Clauses (SCCs)**: EU-approved contract terms ensuring GDPR-level protection
- **Adequacy Decisions**: Transfers only to countries deemed adequate by the European Commission
- **Your Consent**: Where required by law, we obtain explicit consent for international transfers

### 8.2 Microsoft 365 Platform

As THREDrDB operates within Microsoft Excel and uses Microsoft 365 services, data processing locations are governed by Microsoft's data residency policies. Please refer to Microsoft's Trust Center for details on where your Microsoft 365 data is stored and processed.

---

## 9. Data Security

### 9.1 Technical and Organizational Measures

We implement appropriate security measures to protect your data:

**Technical Safeguards:**
- ✅ **Encryption in Transit**: All data transmitted over networks uses TLS 1.2 or higher encryption
- ✅ **Client-Side Processing**: Excel data is processed locally in your browser, never transmitted to our servers
- ✅ **Secure Authentication**: Microsoft Authentication Library (MSAL) for secure identity verification
- ✅ **Code Security**: Regular security updates and dependency vulnerability scanning
- ✅ **Access Controls**: Multi-factor authentication (MFA) required for all administrative access to our systems

**Organizational Safeguards:**
- ✅ **Data Minimization**: We collect only the minimum data necessary (User OID only)
- ✅ **Privacy by Design**: Privacy considerations integrated into all development processes
- ✅ **Security Training**: Regular security awareness training for all personnel
- ✅ **Incident Response Plan**: Documented procedures for handling security incidents

### 9.2 Your Responsibilities

As a user, you are responsible for:
- Keeping your Microsoft 365 account credentials secure
- Protecting your Excel files and data within your local environment
- Using strong passwords and enabling MFA on your Microsoft account
- Reporting any suspected security issues to support@thredr.com

---

## 10. Your Rights Under GDPR

As a data subject, you have the following rights regarding your personal data:

### 10.1 Right to Be Informed
You have the right to clear information about how we process your data (this privacy policy).

### 10.2 Right of Access (Subject Access Request - SAR)
You have the right to request:
- Confirmation that we are processing your personal data
- Access to your personal data
- Additional information about how we process your data

**How to Exercise**: Email support@thredr.com with "Subject Access Request" in the subject line.  
**Response Time**: Within 30 days of receiving your request.  
**Important Note**: Since THREDrDB does not persistently store your User OID or any personal data, we may not have data to provide beyond confirming your subscription status with our payment provider.

### 10.3 Right to Rectification
You have the right to have inaccurate personal data corrected.

**How to Exercise**: Email support@thredr.com with the data you wish to correct.  
**Note**: Your User Object ID is assigned by Microsoft and cannot be changed by us.

### 10.4 Right to Erasure ("Right to Be Forgotten")
You have the right to request deletion of your personal data in certain circumstances:
- You withdraw consent
- Data is no longer necessary for the purposes collected
- You object to processing and there are no overriding legitimate grounds
- Data has been unlawfully processed

**How to Exercise**: Email support@thredr.com with "Data Deletion Request" in the subject line.  
**Response Time**: Within 30 days.  
**Important Note**: Deleting your data will terminate your subscription and access to THREDrDB.

### 10.5 Right to Restriction of Processing
You have the right to request restriction of processing in certain circumstances.

**How to Exercise**: Email support@thredr.com with your restriction request and reason.

### 10.6 Right to Data Portability
You have the right to receive your personal data in a structured, commonly used, and machine-readable format.

**How to Exercise**: Email support@thredr.com with "Data Portability Request" in the subject line.  
**Note**: Given the minimal data we process (User OID only), data portability may be limited.

### 10.7 Right to Object
You have the right to object to processing based on legitimate interests.

**How to Exercise**: Email support@thredr.com with your objection.  
**Effect**: We will cease processing unless we demonstrate compelling legitimate grounds that override your interests.

### 10.8 Rights Related to Automated Decision-Making and Profiling
THREDrDB does not use automated decision-making or profiling. This right is not applicable.

### 10.9 Right to Withdraw Consent
Where processing is based on consent, you may withdraw consent at any time.

**How to Exercise**: Uninstall the THREDrDB add-in or email support@thredr.com.  
**Effect**: Withdrawal does not affect the lawfulness of processing before withdrawal.

### 10.10 Right to Lodge a Complaint
You have the right to lodge a complaint with a supervisory authority if you believe your data protection rights have been violated.

**EU Data Protection Authorities**: https://edpb.europa.eu/about-edpb/board/members_en  
**UK Information Commissioner's Office (ICO)**: https://ico.org.uk/  
**US State Authorities**: Contact your state attorney general or consumer protection office

---

## 11. Children's Privacy

THREDrDB is not intended for use by individuals under the age of 16. We do not knowingly collect personal data from children under 16. If you are a parent or guardian and believe your child has provided personal data to us, please contact support@thredr.com, and we will take steps to delete such information.

---

## 12. California Privacy Rights (CCPA)

If you are a California resident, you have additional rights under the California Consumer Privacy Act (CCPA):

### 12.1 Right to Know
You have the right to request information about the personal information we collect, use, disclose, and sell.

### 12.2 Right to Delete
You have the right to request deletion of your personal information (subject to certain exceptions).

### 12.3 Right to Opt-Out of Sale
We do **not sell** your personal information. This right is not applicable.

### 12.4 Right to Non-Discrimination
We will not discriminate against you for exercising your CCPA rights.

**How to Exercise CCPA Rights**: Email support@thredr.com or call [phone number].  
**Response Time**: Within 45 days of receiving your verifiable request.

---

## 13. Changes to This Privacy Policy

### 13.1 Updates and Notifications

We may update this privacy policy from time to time to reflect:
- Changes in our data processing practices
- New features or functionality
- Legal or regulatory requirements
- Feedback from users or regulators

### 13.2 How We Notify You

When we make material changes to this privacy policy, we will:
- Update the "Last Updated" date at the top of this policy
- Notify you via email (if we have your email address)
- Display a prominent notice within the THREDrDB add-in
- For significant changes: Obtain your renewed consent where required by law

### 13.3 Your Responsibility

We encourage you to review this privacy policy periodically. Your continued use of THREDrDB after changes are posted constitutes your acceptance of the updated policy.

---

## 14. Contact Us

### 14.1 Privacy Questions and Requests

For any privacy-related questions, concerns, or to exercise your data subject rights:

**Email**: support@thredr.com  
**Subject Line**: Use appropriate subject lines like "Privacy Inquiry," "Subject Access Request," "Data Deletion Request," etc.  
**Response Time**: We will respond to all requests within 30 days (or as required by applicable law).

### 14.2 Data Protection Officer

[If you have a DPO, include their contact information here. If not, remove this section.]

---

## 15. Legal Information

### 15.1 Governing Law

This privacy policy and all data processing activities are governed by:
- General Data Protection Regulation (GDPR) - for EU/EEA users
- California Consumer Privacy Act (CCPA) - for California residents
- Other applicable data protection laws based on your location

### 15.2 Compliance Frameworks

THREDrDB is committed to achieving compliance with:
- Microsoft 365 Certification framework
- GDPR requirements
- Industry best practices for data protection

---

## 16. Definitions

**Personal Data**: Any information relating to an identified or identifiable natural person.

**User Object ID (OID)**: A unique pseudonymous identifier assigned by Microsoft to your Microsoft 365 account. It is a GUID (Globally Unique Identifier) that does not directly reveal your identity.

**Pseudonymization**: Processing of personal data in such a way that it can no longer be attributed to a specific data subject without additional information.

**Data Subject**: An individual whose personal data is being processed.

**Data Controller**: The entity that determines the purposes and means of processing personal data (Thredr LLC for THREDrDB).

**Data Processor**: An entity that processes personal data on behalf of the data controller (e.g., our subscription service provider).

---

## 17. Acknowledgment

By installing and using THREDrDB, you acknowledge that you have read, understood, and agree to be bound by this Privacy Policy.

---

**End of Privacy Policy**

Last Reviewed: November 3, 2025  
Version: 2.0 (GDPR-Enhanced)
