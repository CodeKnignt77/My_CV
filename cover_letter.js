const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');
const fs = require('fs');
const docxConverter = require('docx-pdf');

const doc = new Document({
  styles: {
    default: { 
      document: { 
        run: { font: "Arial", size: 22 } // 11pt
      } 
    }
  },
  sections: [{
    properties: {
      page: {
        size: {
          width: 12240,   // US Letter
          height: 15840
        },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } // 1 inch margins
      }
    },
    children: [
      // Header with contact info
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 120 },
        children: [
          new TextRun({
            text: "Yasser Chaouki",
            bold: true,
            size: 24
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({
            text: "+212 707512723"
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({
            text: "chaouki0yasser5@gmail.com"
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({
            text: "https://www.linkedin.com/in/yasser-chaouki"
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 240 },
        children: [
          new TextRun({
            text: "Taroudant, Morocco"
          })
        ]
      }),

      // Date
      new Paragraph({
        spacing: { after: 240 },
        children: [
          new TextRun({
            text: "January 28, 2026"
          })
        ]
      }),

      // Recipient
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({
            text: "Hyphen Cyber",
            bold: true
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({
            text: "Recruitment team"
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 240 },
        children: [
          new TextRun({
            text: "Sharjah, United Arab Emirates"
          })
        ]
      }),

      // Subject line
      new Paragraph({
        spacing: { after: 240 },
        children: [
          new TextRun({
            text: "Re: Cybersecurity professional opportunity",
            bold: true
          })
        ]
      }),

      // Opening
      new Paragraph({
        spacing: { after: 180 },
        children: [
          new TextRun({
            text: "Dear Hyphen Cyber recruitment team,"
          })
        ]
      }),

      // Body paragraph 1
      new Paragraph({
        spacing: { after: 180 },
        children: [
          new TextRun({
            text: "I am writing to express my strong interest in cybersecurity opportunities with Hyphen Cyber. As a first-year cybersecurity engineering student at ENSIASD with a robust foundation in cryptography, algorithm design, and mathematical analysis, I am eager to contribute to your mission of connecting high-caliber cybersecurity professionals across industries. Your commitment to magnetizing global connections and fostering growth within the cybersecurity community aligns perfectly with my professional aspirations and values."
          })
        ]
      }),

      // Body paragraph 2
      new Paragraph({
        spacing: { after: 180 },
        children: [
          new TextRun({
            text: "My background demonstrates a deep commitment to cybersecurity through both academic rigor and hands-on research. During my two years in the prestigious CPGE program at Moulay Al Hassan, I developed advanced mathematical foundations in linear algebra, multivariable calculus, and discrete mathematics that directly inform my approach to security problems. I subsequently conducted comprehensive cryptographic research on RSA public-key cryptosystems, where I analyzed the computational complexity of integer factorization and designed lattice-based side-channel attacks using Graham-Schmidt orthogonalization. This research, which I documented and published on LinkedIn, demonstrates my ability to bridge theoretical number theory with practical information security applications."
          })
        ]
      }),

      // Body paragraph 3
      new Paragraph({
        spacing: { after: 180 },
        children: [
          new TextRun({
            text: "Currently, I am expanding my technical capabilities through multiple active projects. I am engineering a Retrieval-Augmented Generation system optimized for symbolic mathematics using LangChain and Pinecone, achieving measurable improvements in reasoning accuracy. Additionally, I am leading the development of a hyperlocal delivery infrastructure at Enactus ENSIASD, where I architect backend systems for real-time order processing while managing the full software development lifecycle. I am also actively pursuing the ISC2 certified in cybersecurity credential and continuously building my skills through Microsoft's ML-For-Beginners repository, demonstrating my commitment to autonomous learning and professional development."
          })
        ]
      }),

      // Body paragraph 4
      new Paragraph({
        spacing: { after: 180 },
        children: [
          new TextRun({
            text: "My technical expertise spans security and cryptography, including RSA cryptanalysis, integer factorization, and network security protocols such as TCP/IP, DNS, and VPNs. I am proficient in Python, Java, SQL, and C/C++, with growing capabilities in machine learning frameworks including PyTorch, TensorFlow, and Scikit-learn. I have achieved 92% accuracy in classifying network intrusions using predictive models, demonstrating my ability to apply machine learning to security challenges. My recognition as a national finalist in the Math&Maroc Competition 2024 further validates my problem-solving capabilities and analytical rigor."
          })
        ]
      }),

      // Body paragraph 5
      new Paragraph({
        spacing: { after: 180 },
        children: [
          new TextRun({
            text: "I am particularly drawn to Hyphen Cyber's focus on connecting talented professionals with organizations that value cybersecurity excellence. I believe my combination of theoretical knowledge, research experience, and practical implementation skills positions me as a valuable candidate for opportunities within your network. I am eager to contribute my analytical mindset, mathematical rigor, and technical capabilities to help organizations strengthen their security posture against evolving cyber threats."
          })
        ]
      }),

      // Closing paragraph
      new Paragraph({
        spacing: { after: 180 },
        children: [
          new TextRun({
            text: "I would welcome the opportunity to discuss how my background, skills, and passion for cybersecurity align with opportunities in your network. Thank you for considering my application. I look forward to the possibility of connecting with you and exploring how I can contribute to the cybersecurity community through Hyphen Cyber."
          })
        ]
      }),

      // Sign off
      new Paragraph({
        spacing: { before: 180, after: 60 },
        children: [
          new TextRun({
            text: "Sincerely,"
          })
        ]
      }),

      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({
            text: "Yasser Chaouki",
            bold: true
          })
        ]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  const outputDir = "c:\\Users\\Lenovo\\Downloads\\CV_ATS\\output";
  
  // Create output directory if it doesn't exist
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }
  
  const docxPath = `${outputDir}\\Yasser_Chaouki_Cover_Letter_Hyphen_Cyber.docx`;
  const pdfPath = `${outputDir}\\Yasser_Chaouki_Cover_Letter_Hyphen_Cyber.pdf`;
  
  // Save DOCX
  fs.writeFileSync(docxPath, buffer);
  console.log("DOCX created successfully!");
  
  // Convert to PDF
  docxConverter(docxPath, pdfPath, (err, result) => {
    if (err) {
      console.error("Error converting to PDF:", err);
    } else {
      console.log("PDF created successfully at: " + pdfPath);
    }
  });
});