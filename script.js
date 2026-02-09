// Global Variables
let employees = [];
let currentIndex = 0;
let validEmployees = [];

// Initialize
document.addEventListener("DOMContentLoaded", () => {
  document
    .getElementById("excelFile")
    .addEventListener("change", handleFileUpload);
  showStatus("يرجى تحميل ملف Excel للبدء", "info");
});

// Handle File Upload
function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  showStatus("جاري تحميل الملف...", "info");

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      employees = XLSX.utils.sheet_to_json(firstSheet);

      console.log("Sample employee data:", employees[0]); // Debug
      console.log("Available columns:", Object.keys(employees[0])); // Debug

      // Filter valid employees - check for Name column
      validEmployees = employees.filter((emp) => {
        const hasName =
          emp["Name"] &&
          emp["Name"].trim() !== "" &&
          !emp["Name"].includes("Total") &&
          emp["Name"] !== "Total";
        return hasName;
      });

      if (validEmployees.length === 0) {
        showStatus("لم يتم العثور على موظفين صالحين في الملف", "error");
        return;
      }

      currentIndex = 0;
      displayEmployee(currentIndex);
      showStatus(`✅ تم تحميل ${validEmployees.length} موظف بنجاح`, "success");

      // Show navigation and actions
      document.getElementById("navigationSection").style.display = "flex";
      document.getElementById("actionSection").style.display = "flex";
      updateNavigation();
    } catch (error) {
      console.error("Error:", error);
      showStatus("خطأ في قراءة الملف. يرجى التأكد من صحة الملف", "error");
    }
  };

  reader.onerror = () => showStatus("خطأ في قراءة الملف", "error");
  reader.readAsArrayBuffer(file);
}

// Display Employee Form
function displayEmployee(index) {
  if (!validEmployees[index]) return;

  const employee = validEmployees[index];
  const formHTML = generateGovernmentForm(employee);
  document.getElementById("formContainer").innerHTML = formHTML;
  updateNavigation();
}

// Generate Exact Government Form
function generateGovernmentForm(employee) {
  // Extract data from Excel columns
  const arabicNames = extractArabicNames(employee);

  // Get National ID from the correct column
  const nationalId =
    employee["الرقم الوطني/ جواز السفر"] ||
    employee["ID"] ||
    employee["DME"] ||
    "";

  const taxYear = "2025";

  return `
        <div class="government-form" id="currentForm">
            <!-- Header Section -->
            <div class="form-header">
                <img src="crown.png" alt="شعار المملكة" class="form-logo">
                <div class="form-title-main">المملكـــة الأردنيــة الهاشميـــة</div>
                <div class="form-title-secondary">شهادة مجموع الرواتب والأجور والضريبة المقتطعة</div>
                <div class="form-subtitle">
                    استناداً لأحكام الفقرة (أ) من المادة السادسة للتعليمات رقم ( 1 ) لسنة 2015
                </div>
                <div class="form-subtitle">
                    والمعدلة بالأستناد لأحكام الفقرة ( و ) من المادة ( 12 ) من قانون ضريبة الدخل رقم ( 34 ) لسنة 2014 م.
                </div>
            </div>

            <!-- Employee Information Table -->
            <table class="govt-table">
                <tr>
                    <th colspan="4" class="section-header-cell">معلومات الموظف</th>
                </tr>
                <tr>
                    <th colspan="4" class="section-header-cell">اسم الموظف</th>
                </tr>
                <tr>
                    <th class="label-cell">الاسم</th>
                    <th class="label-cell">الأب</th>
                    <th class="label-cell">الجد</th>
                    <th class="label-cell">العائله</th>
                </tr>
                <tr>
                    <td class="value-cell">${arabicNames.firstName}</td>
                    <td class="value-cell">${arabicNames.fatherName}</td>
                    <td class="value-cell">${arabicNames.grandFatherName}</td>
                    <td class="value-cell">${arabicNames.familyName}</td>
                </tr>
                <tr class="empty-row">
                    <td colspan="4"></td>
                </tr>
                <tr>
                    <th class="label-cell">الرقم الضريبي</th>
                    <th class="label-cell" colspan="2">الرقم الوطني/ جواز السفر</th>
                    <th class="label-cell">الرمز البريدي</th>
                </tr>
                <tr>
                    <td class="value-cell">-</td>
                    <td class="value-cell" colspan="2">${nationalId}</td>
                    <td class="value-cell">-</td>
                </tr>
                <tr>
                    <th class="label-cell" colspan="2">العنـــــوان</th>
                    <th class="label-cell" colspan="2">الهاتف</th>
                </tr>
                <tr>
                    <td class="value-cell" colspan="2">عمان</td>
                    <td class="value-cell" colspan="2">-</td>
                </tr>
                <tr>
                    <th class="label-cell">الفترة الضريبية</th>
                    <th class="label-cell" colspan="2">مدة العمل لغاية الفترة الضريبية</th>
                    <th class="label-cell">تاريخ انتهاء العمل (الإنهاء الفعلي)</th>
                </tr>
                <tr>
                    <td class="value-cell">${taxYear}</td>
                    <td class="value-cell" colspan="2"></td>
                    <td class="value-cell">-</td>
                </tr>
            </table>

            <!-- Financial Information Table -->
            <table class="financial-table">
                <tr>
                    <th rowspan="2" class="section-header-cell">إجمالي الرواتب والأجور</th>
                    <th colspan="2">القيمة</th>
                    <th rowspan="2" class="section-header-cell">الضريبة المقتطعة من إجمالي</th>
                    <th colspan="2">القيمة</th>
                </tr>
                <tr>
                    <th>دينار</th>
                    <th>فلس</th>
                    <th>دينار</th>
                    <th>فلس</th>
                </tr>
                <tr>
                    <td class="col-label">الرواتب والأجور</td>
                    <td class="value-cell">-</td>
                    <td class="value-cell">-</td>
                    <td class="col-label">الرواتب والأجور</td>
                    <td class="value-cell">-</td>
                    <td class="value-cell">-</td>
                </tr>
                <tr>
                    <td class="col-label">الرواتب والأجور غير الشهرية</td>
                    <td class="value-cell">-</td>
                    <td class="value-cell">-</td>
                    <td class="col-label">الرواتب والأجور غير الشهرية</td>
                    <td class="value-cell">-</td>
                    <td class="value-cell">-</td>
                </tr>
                <tr>
                    <td class="col-label">مكافآت أعضاء مجلس الإدارة</td>
                    <td class="value-cell">-</td>
                    <td class="value-cell">-</td>
                    <td class="col-label">مكافآت أعضاء مجلس الإدارة</td>
                    <td class="value-cell">-</td>
                    <td class="value-cell">-</td>
                </tr>
                <tr>
                    <td class="col-label">مكافأة نهاية الخدمة</td>
                    <td class="value-cell">-</td>
                    <td class="value-cell">-</td>
                    <td class="col-label">مكافأة نهاية الخدمة</td>
                    <td class="value-cell">-</td>
                    <td class="value-cell">-</td>
                </tr>
                <tr>
                    <td class="col-label">أي مبالغ أخرى</td>
                    <td class="value-cell">-</td>
                    <td class="value-cell">-</td>
                    <td class="col-label">أي مبالغ أخرى</td>
                    <td class="value-cell">-</td>
                    <td class="value-cell">-</td>
                </tr>
                <tr style="background: #f5f5f5;">
                    <td class="col-label"><strong>المجمــــــــوع</strong></td>
                    <td class="value-cell">-</td>
                    <td class="value-cell">-</td>
                    <td class="col-label"><strong>المجمــــــــوع</strong></td>
                    <td class="value-cell">-</td>
                    <td class="value-cell">-</td>
                </tr>
            </table>

            <!-- Declaration -->
            <div class="declaration">
                أشهد أن المعلومات المذكورة أعلاه صحيحة ودقيقة وغير منقوصة وأنني قمت بتبليغ ضريبة الدخل المقتطعة والمبينة أعلاه إلى دائرة ضريبة الدخل والمبيعات .
            </div>

            <!-- Company Information -->
            <table class="company-table">
                <tr>
                    <td class="company-label-cell">اسم صاحب العمل</td>
                    <td colspan="2" class="value-cell">شركة ديلوتيم ميدل إيست</td>
                </tr>
                <tr>
                    <td class="company-label-cell">الرقم الضريبي</td>
                    <td colspan="2" class="value-cell">17934435</td>
                </tr>
                <tr>
                    <td class="company-label-cell" colspan="3">ختم توقيع صاحب العمل</td>
                </tr>
                <tr>
                    <td colspan="3" class="signature-cell"></td>
                </tr>
            </table>

            <div class="date-line">التاريخ : 10/1/2026</div>
        </div>
    `;
}

// Extract Arabic Names from Excel - UPDATED FOR YOUR COLUMNS
function extractArabicNames(employee) {
  console.log("Processing employee:", employee); // Debug

  // Based on your Excel screenshot, the Arabic columns are:
  // Column M: الاسم (First Name)
  // Column N: الأب (Father)
  // Column O: الجد (Grandfather)
  // Column P: العائلة (Family Name)

  // Try to find the Arabic name columns by their header names
  let firstName = "";
  let fatherName = "";
  let grandFatherName = "";
  let familyName = "";

  // Check all possible column name variations
  for (let key in employee) {
    const lowerKey = key.toLowerCase();

    // First name variations
    if (lowerKey.includes("الاسم") && !lowerKey.includes("الرباعي")) {
      firstName = employee[key] || "";
    }
    // Father name variations
    else if (lowerKey.includes("الأب") || lowerKey.includes("اب")) {
      fatherName = employee[key] || "";
    }
    // Grandfather variations
    else if (lowerKey.includes("الجد") || lowerKey.includes("جد")) {
      grandFatherName = employee[key] || "";
    }
    // Family name variations
    else if (lowerKey.includes("العائله") || lowerKey.includes("عائلة")) {
      familyName = employee[key] || "";
    }
  }

  // If we couldn't find the Arabic columns, try to extract from the English Name column
  if (!firstName && employee["Name"]) {
    console.log("Falling back to parsing Name column");
    const fullName = employee["Name"].trim();
    const parts = fullName.split(/\s+/);

    // For names like "Abdallah Sami Abdallah Toughoz"
    // Map to Arabic if available in other fields, otherwise use English
    firstName = parts[0] || "";
    fatherName = parts[1] || "";
    grandFatherName = parts[2] || "";
    familyName = parts.slice(3).join(" ") || "";
  }

  console.log("Extracted names:", {
    firstName,
    fatherName,
    grandFatherName,
    familyName,
  }); // Debug

  return {
    firstName: firstName.trim(),
    fatherName: fatherName.trim(),
    grandFatherName: grandFatherName.trim(),
    familyName: familyName.trim(),
  };
}

// Navigation Functions
function nextEmployee() {
  if (currentIndex < validEmployees.length - 1) {
    currentIndex++;
    displayEmployee(currentIndex);
  }
}

function prevEmployee() {
  if (currentIndex > 0) {
    currentIndex--;
    displayEmployee(currentIndex);
  }
}

function updateNavigation() {
  const counter = document.getElementById("employeeCounter");
  const empName = document.getElementById("employeeName");
  const prevBtn = document.getElementById("prevBtn");
  const nextBtn = document.getElementById("nextBtn");

  if (validEmployees.length > 0) {
    counter.textContent = `${currentIndex + 1} / ${validEmployees.length}`;

    const names = extractArabicNames(validEmployees[currentIndex]);
    const fullName =
      `${names.firstName} ${names.fatherName} ${names.grandFatherName} ${names.familyName}`.trim();
    empName.textContent =
      fullName || validEmployees[currentIndex]["Name"] || "";

    prevBtn.disabled = currentIndex === 0;
    nextBtn.disabled = currentIndex === validEmployees.length - 1;
  }
}

// Status Message
function showStatus(message, type) {
  const statusDiv = document.getElementById("status");
  statusDiv.textContent = message;
  statusDiv.className = `status-message ${type}`;
}

// Download Current PDF
async function downloadCurrentPDF() {
  if (validEmployees.length === 0) {
    alert("لا يوجد موظفين");
    return;
  }

  showStatus("جاري إنشاء PDF...", "info");

  try {
    const employee = validEmployees[currentIndex];
    const names = extractArabicNames(employee);
    const fileName = `tax-form-${names.firstName || "employee"}-${currentIndex + 1}.pdf`;

    await generateSinglePDF("currentForm", fileName);
    showStatus("✅ تم تحميل PDF بنجاح", "success");
  } catch (error) {
    console.error("PDF Error:", error);
    showStatus("❌ حدث خطأ في إنشاء PDF", "error");
  }
}

// Download All PDFs - Enhanced
async function downloadAllPDFs() {
  if (validEmployees.length === 0) {
    alert("لا يوجد موظفين");
    return;
  }

  const confirmed = confirm(
    `هل تريد إنشاء ${validEmployees.length} نموذج PDF؟\nقد يستغرق بعض الوقت...`,
  );
  if (!confirmed) return;

  const modal = document.getElementById("progressModal");
  const progressFill = document.getElementById("progressFill");
  const progressText = document.getElementById("progressText");

  modal.classList.add("active");

  try {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a4",
      compress: true,
    });

    let firstPage = true;

    for (let i = 0; i < validEmployees.length; i++) {
      // Update progress
      const progress = Math.round(((i + 1) / validEmployees.length) * 100);
      progressFill.style.width = progress + "%";
      progressText.textContent = `${i + 1} من ${validEmployees.length}`;

      // Create temporary form
      const tempDiv = document.createElement("div");
      tempDiv.innerHTML = generateGovernmentForm(validEmployees[i]);
      tempDiv.style.position = "absolute";
      tempDiv.style.left = "-9999px";
      tempDiv.style.width = "900px";
      tempDiv.style.background = "white";
      document.body.appendChild(tempDiv);

      // Wait for images to load
      await waitForImages(tempDiv);

      // Generate canvas
      const canvas = await html2canvas(
        tempDiv.querySelector(".government-form"),
        {
          scale: 2,
          useCORS: true,
          allowTaint: true,
          logging: false,
          backgroundColor: "#ffffff",
          width: 900,
          windowWidth: 900,
        },
      );

      // Add to PDF
      const imgData = canvas.toDataURL("image/jpeg", 0.95);
      const imgWidth = 210; // A4 width
      const imgHeight = (canvas.height * imgWidth) / canvas.width;

      if (!firstPage) {
        pdf.addPage();
      }
      firstPage = false;

      pdf.addImage(imgData, "JPEG", 0, 0, imgWidth, imgHeight);

      // Cleanup
      document.body.removeChild(tempDiv);

      // Small delay to prevent browser freeze
      await sleep(50);
    }

    // Save PDF
    pdf.save(`employee-tax-forms-all-${validEmployees.length}.pdf`);

    modal.classList.remove("active");
    showStatus(`✅ تم إنشاء ${validEmployees.length} نموذج بنجاح`, "success");
  } catch (error) {
    console.error("Batch PDF Error:", error);
    modal.classList.remove("active");
    showStatus("❌ حدث خطأ في إنشاء الملفات", "error");
  }
}

// Generate Single PDF
async function generateSinglePDF(elementId, fileName) {
  const element = document.getElementById(elementId);
  if (!element) throw new Error("Element not found");

  const canvas = await html2canvas(element, {
    scale: 2,
    useCORS: true,
    allowTaint: true,
    logging: false,
    backgroundColor: "#ffffff",
  });

  const imgData = canvas.toDataURL("image/jpeg", 0.95);
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({
    orientation: "portrait",
    unit: "mm",
    format: "a4",
  });

  const imgWidth = 210;
  const imgHeight = (canvas.height * imgWidth) / canvas.width;

  pdf.addImage(imgData, "JPEG", 0, 0, imgWidth, imgHeight);
  pdf.save(fileName);
}

// Helper: Wait for images to load
function waitForImages(element) {
  return new Promise((resolve) => {
    const images = element.querySelectorAll("img");
    if (images.length === 0) {
      resolve();
      return;
    }

    let loadedCount = 0;
    images.forEach((img) => {
      if (img.complete) {
        loadedCount++;
      } else {
        img.onload = () => {
          loadedCount++;
          if (loadedCount === images.length) resolve();
        };
        img.onerror = () => {
          loadedCount++;
          if (loadedCount === images.length) resolve();
        };
      }
    });

    if (loadedCount === images.length) resolve();
  });
}

// Helper: Sleep function
function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
