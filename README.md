# 📝 DM Proposal Generator

[![Live Demo](https://img.shields.io/badge/🌐%20Live-DM%20Proposal%20Generator-brightgreen)](https://dm-proposal-generator-566552386634.us-central1.run.app/)

## 📌 Overview

**DM Proposal Generator** is an AI-powered tool designed to automate the creation of digital marketing proposals. It streamlines proposal drafting using customizable templates, ensuring professionalism and efficiency.

---

## 🚀 Features

- 📄 **Automated Proposal Creation** – Generates marketing proposals based on user-provided details.
- 🧩 **Customizable Templates** – Offers predefined templates for different marketing strategies.
- ✨ **Clean Formatting** – Ensures uniform structure and easy readability.
- 🧠 **Smart Suggestions** – Helps optimize content for clarity and professionalism.

---

## 🛠️ Installation

To run the project locally:

```bash
git clone https://github.com/Aswinthmani2003/DM-Proposal-Generator.git
cd DM-Proposal-Generator
```

Create and activate a virtual environment:

- On **macOS/Linux**:

```bash
python3 -m venv venv
source venv/bin/activate
```

- On **Windows**:

```bash
python -m venv venv
venv\Scripts\activate
```

Install the required dependencies:

```bash
pip install -r requirements.txt
```

---

## 💡 Usage

1. **Run the app**:
   ```bash
   python app.py
   ```

2. **Input Details**:
   Provide client information, services required, objectives, etc.

3. **Choose a Template**:
   Select from various marketing combinations like:
   - SEO
   - SMM (Social Media Marketing)
   - Google Ads
   - Meta Ads
   - Email Marketing

4. **Generate & Review**:
   Get your draft proposal, review it, and edit if needed.

5. **Export**:
   Save the proposal as a PDF or Word document.

---

## 🐳 Docker Deployment

You can also deploy the app with Docker:

```bash
docker build -t dm-proposal-generator .
docker run -p 5000:5000 dm-proposal-generator
```

Open in browser at: [http://localhost:5000](http://localhost:5000)

---

## 📁 Proposal Templates Included

- DM Proposal - All.docx
- Only Email Marketing.docx
- Only Google Ads Campaign.docx
- Only Meta Ads Campaign.docx
- Only SEO.docx
- Only SMM.docx
- SEO & Google Ads Campaign.docx
- SMM & Meta Ads Campaigns.docx
- SMM, Google Ads & Meta Ads Campaigns.docx
- SMM, Meta & Google Ads and SEO.docx

---

## 🤝 Contributing

1. Fork this repo
2. Create a new branch (`git checkout -b feature-name`)
3. Commit changes (`git commit -m "Added feature"`)
4. Push to your branch (`git push origin feature-name`)
5. Open a Pull Request

---

## 📞 Contact

- 📧 Email: aswinthmani10@gmail.com  
- 🐙 GitHub: [@Aswinthmani2003](https://github.com/Aswinthmani2003)

---

## ☁️ Deployment

The app is currently deployed and live on Google Cloud Platform:

🌐 **Live Demo**: [https://dm-proposal-generator-566552386634.us-central1.run.app/](https://dm-proposal-generator-566552386634.us-central1.run.app/)

To deploy on your own GCP project, include an `app.yaml`:

```yaml
runtime: python310
entrypoint: streamlit run app.py --server.port=8080 --server.enableCORS=false

instance_class: F2
automatic_scaling:
  target_cpu_utilization: 0.65
  min_instances: 1
  max_instances: 2
```

Deploy using:

```bash
gcloud app deploy
```

---

> Built with ❤️ to simplify digital marketing proposals.
