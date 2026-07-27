# OMKAR RETAIL VENTURES Website

Modern, mobile-first recruitment website built for:
- Job seekers
- Picker & Packer candidates
- Store staff
- Inbound executive candidates
- Zepto Dark Store candidates (OMKAR internal hiring)

Domain target: `omkarretailventures.in`

## Pages

- `index.html` - Home + About company information
- `apply.html` - Careers page with roles, eligibility, hiring details, and application form
- `employee.html` - Employee support, onboarding, and culture page
- `contact.html` - Contact + support + map + social + WhatsApp
- `admin.html` - Admin dashboard (applications, Zepto leads, contacts, support)

## Forms & Data

Forms are connected in `forms.js` with:
- Validation
- Honeypot anti-spam field
- Basic fast-submit bot blocking
- Success/error states
- Local storage fallback

If Firebase config is provided in `config.js`, forms will also:
- Save records to Firestore
- Upload resumes to Firebase Storage
- Show data in Admin dashboard from Firebase

## Configure Before Launch

Edit `config.js`:
- `phone`
- `whatsappNumber`
- `email`
- `officeAddress`
- `social` links
- `adminPin`
- `analyticsMeasurementId` (Google Analytics)
- `web3formsAccessKey` (optional email notifications)
- `firebase` object (for database + resume storage)

## Firebase Setup (Recommended)

1. Create Firebase project.
2. Enable Firestore Database.
3. Enable Storage.
4. Get web app config and paste in `config.js -> firebase`.
5. Optional but recommended: enable Firebase Auth and protect admin access in production.

Suggested Firestore collections used by this site:
- `applications`
- `employerLeads`
- `contacts`
- `supports`

## GitHub Pages Deployment

```powershell
cd "C:\Users\Rajeshwari\Desktop\docs"
git init
git add .
git commit -m "Launch OMKAR RETAIL VENTURES website"
git branch -M main
git remote add origin https://github.com/<your-username>/<repo-name>.git
git push -u origin main
```

Then on GitHub:
1. Repo -> Settings -> Pages
2. Source: Deploy from a branch
3. Branch: `main` and `/ (root)`

`CNAME` file is already added with `omkarretailventures.in`.

## GoDaddy DNS (for GitHub Pages)

Add these records:
- `A` -> Host `@` -> `185.199.108.153`
- `A` -> Host `@` -> `185.199.109.153`
- `A` -> Host `@` -> `185.199.110.153`
- `A` -> Host `@` -> `185.199.111.153`
- `CNAME` -> Host `www` -> `<your-username>.github.io`

After propagation, enable HTTPS in GitHub Pages settings.

## SEO Files Included

- `robots.txt`
- `sitemap.xml`
- per-page meta description and canonical tags

## Notes

- `admin.html` PIN protection is basic and suitable only as a starter.
- For strong security, use Firebase Auth + Firestore security rules.
- WhatsApp QR and brand image are in `assets/images/`.


