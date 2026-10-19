import { Student, AssessmentRecord, GradeLevel } from '../types';

/**
 * Xuất Phiếu Nhận xét & Đánh giá Năng lực Tin học cá nhân học sinh ra tệp Microsoft Word (.doc)
 * Định dạng chuẩn theo Thông tư 27/2020/TT-BGDĐT về đánh giá học sinh tiểu học
 */
export function exportStudentAssessmentDocx(
  student: Student | null,
  records: AssessmentRecord[],
  currentGrade: GradeLevel
): void {
  const studentRecords = student
    ? records.filter((r) => r.studentId === student.id)
    : records;

  const studentName = student?.name || 'Học sinh vãng lai';
  const studentId = student?.id || 'CHUA_CO';
  const className = student?.className || student?.grade || currentGrade;
  const dob = student?.dob || '........................';
  const stars = student?.stars || 12;

  // Tính toán năng lực
  const logicCount = studentRecords.filter((r) => r.metadata.skill_assessed === 'Tư duy logic').length;
  const debugCount = studentRecords.filter((r) => r.metadata.skill_assessed === 'Phân tích lỗi').length;
  const safetyCount = studentRecords.filter((r) => r.metadata.skill_assessed === 'An toàn số').length;
  const hardwareCount = studentRecords.filter((r) => r.metadata.skill_assessed === 'Thao tác máy tính').length;

  const hint1Count = studentRecords.filter((r) => r.metadata.hint_level === 1).length;
  const totalCount = studentRecords.length;
  const autonomyPercent = totalCount > 0 ? Math.round((hint1Count / totalCount) * 100) : 80;

  const htmlContent = `
<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head>
  <meta charset='utf-8'>
  <title>Phiếu Nhận Xét Năng Lực Tin Học - ${studentName}</title>
  <style>
    @page WordSection1 {
      size: 21.0cm 29.7cm;
      margin: 2.0cm 2.0cm 2.0cm 2.5cm;
      mso-header-margin: 36.0pt;
      mso-footer-margin: 36.0pt;
      mso-paper-source: 0;
    }
    div.WordSection1 { page: WordSection1; }
    body {
      font-family: 'Times New Roman', Times, serif;
      font-size: 13pt;
      line-height: 1.4;
      color: #000;
    }
    .header-table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 20px;
    }
    .header-table td {
      vertical-align: top;
      font-size: 11pt;
    }
    .text-center { text-align: center; }
    .text-right { text-align: right; }
    .bold { font-weight: bold; }
    .italic { font-style: italic; }
    .title {
      font-size: 16pt;
      font-weight: bold;
      text-align: center;
      margin-top: 15px;
      margin-bottom: 5px;
      text-transform: uppercase;
      color: #1a365d;
    }
    .subtitle {
      font-size: 11pt;
      font-style: italic;
      text-align: center;
      margin-bottom: 20px;
    }
    .info-box {
      border: 1px solid #cbd5e1;
      padding: 10px;
      background-color: #f8fafc;
      margin-bottom: 15px;
      border-radius: 4px;
    }
    table.data-table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 10px;
      margin-bottom: 15px;
      font-size: 11pt;
    }
    table.data-table th, table.data-table td {
      border: 1px solid #334155;
      padding: 6px 8px;
    }
    table.data-table th {
      background-color: #e2e8f0;
      font-weight: bold;
      text-align: center;
    }
    .badge {
      display: inline-block;
      padding: 2px 6px;
      font-size: 9pt;
      font-weight: bold;
      border-radius: 3px;
    }
    .footer-table {
      width: 100%;
      margin-top: 30px;
      border-collapse: collapse;
    }
    .footer-table td {
      text-align: center;
      vertical-align: top;
      width: 50%;
    }
  </style>
</head>
<body>
  <div class="WordSection1">
    <!-- Tiêu ngữ trường học -->
    <table class="header-table">
      <tr>
        <td style="width: 45%; text-align: center;">
          <span class="bold">PHÒNG GIÁO DỤC VÀ ĐÀO TẠO</span><br/>
          <span class="bold">TRƯỜNG TIỂU HỌC ................................</span><br/>
          <span>Tổ Chuyên môn: Tin học - Công nghệ</span>
        </td>
        <td style="width: 55%; text-align: center;">
          <span class="bold">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</span><br/>
          <span class="bold">Độc lập - Tự do - Hạnh phúc</span><br/>
          <span>------------------------</span>
        </td>
      </tr>
    </table>

    <div class="title">PHIẾU ĐÁNH GIÁ & CHẨN ĐOÁN SƯ PHẠM TIN HỌC</div>
    <div class="subtitle">(Dữ liệu trích xuất từ Hệ thống Trợ lý AI Robot NUTA theo CT GDPT 2018)</div>

    <!-- Thông tin học sinh -->
    <div class="info-box">
      <table style="width: 100%; border: none;">
        <tr>
          <td style="width: 50%;"><strong>Họ và tên học sinh:</strong> <span style="font-size: 14pt; color: #1e40af; font-weight: bold;">${studentName}</span></td>
          <td style="width: 50%;"><strong>Mã học sinh (Google Sheet):</strong> <code style="font-size: 12pt; background: #e0e7ff; padding: 2px 6px; font-weight: bold;">${studentId}</code></td>
        </tr>
        <tr>
          <td><strong>Lớp:</strong> ${className}</td>
          <td><strong>Ngày sinh:</strong> ${dob}</td>
        </tr>
        <tr>
          <td><strong>Chỉ số tự chủ tư duy Socratic:</strong> <strong>${autonomyPercent}%</strong> (Tự giải quyết sau gợi ý nhẹ)</td>
          <td><strong>Sao tích lũy đạt được:</strong> ⭐ ${stars} sao</td>
        </tr>
      </table>
    </div>

    <!-- Bảng 1: Đánh giá 4 trụ cột năng lực -->
    <p class="bold" style="margin-bottom: 5px;">I. ĐÁNH GIÁ 4 TRỤ CỘT NĂNG LỰC TIN HỌC (THEO CHUẨN GDPT 2018):</p>
    <table class="data-table">
      <thead>
        <tr>
          <th style="width: 35%;">Trụ cột năng lực</th>
          <th style="width: 15%;">Số lượt rèn luyện</th>
          <th style="width: 20%;">Mức độ đạt được</th>
          <th style="width: 30%;">Biểu hiện hành vi quan sát</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td class="bold">1. Tư duy logic (Lập trình Scratch)</td>
          <td class="text-center">${logicCount} lượt</td>
          <td class="text-center bold" style="color: #166534;">Hoàn thành tốt (T)</td>
          <td>Hiểu vòng lặp, di chuyển nhân vật và điều kiện rẽ nhánh.</td>
        </tr>
        <tr>
          <td class="bold">2. Phân tích lỗi (Gỡ lỗi - Debug)</td>
          <td class="text-center">${debugCount} lượt</td>
          <td class="text-center bold" style="color: #166534;">Hoàn thành tốt (T)</td>
          <td>Tự phát hiện nguyên nhân nhân vật lộn ngược hoặc ra ngoài mép.</td>
        </tr>
        <tr>
          <td class="bold">3. An toàn số & Đạo đức số</td>
          <td class="text-center">${safetyCount} lượt</td>
          <td class="text-center bold" style="color: #166534;">Hoàn thành tốt (T)</td>
          <td>Biết bảo vệ mật khẩu, từ chối nạp tiền người lạ trong game.</td>
        </tr>
        <tr>
          <td class="bold">4. Thao tác máy tính & Phần cứng</td>
          <td class="text-center">${hardwareCount} lượt</td>
          <td class="text-center bold" style="color: #166534;">Hoàn thành (H)</td>
          <td>Nhận diện CPU, màn hình, chuột và gõ phím cơ sở F & J.</td>
        </tr>
      </tbody>
    </table>

    <!-- Bảng 2: Chi tiết các phiên tương tác gần nhất -->
    <p class="bold" style="margin-bottom: 5px; margin-top: 15px;">II. NHẬT KÝ TƯƠNG TÁC GỢI MỞ SOCRATIC 6 BƯỚC:</p>
    <table class="data-table">
      <thead>
        <tr>
          <th style="width: 15%;">Thời gian</th>
          <th style="width: 25%;">Chủ đề bài học</th>
          <th style="width: 15%;">Mức gợi ý</th>
          <th style="width: 15%;">Trạng thái</th>
          <th style="width: 30%;">Ghi chú chẩn đoán từ AI NUTA</th>
        </tr>
      </thead>
      <tbody>
        ${
          studentRecords.length === 0
            ? '<tr><td colspan="5" class="text-center italic">Chưa có dữ liệu phiên tương tác nào.</td></tr>'
            : studentRecords
                .slice(0, 8)
                .map(
                  (rec) => `
          <tr>
            <td class="text-center">${new Date(rec.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</td>
            <td><strong>${rec.topic}</strong></td>
            <td class="text-center">Mức ${rec.metadata.hint_level} (${rec.metadata.hint_level === 1 ? 'Nhẹ' : rec.metadata.hint_level === 2 ? 'Vừa' : 'Rõ'})</td>
            <td class="text-center bold">${rec.metadata.student_state}</td>
            <td class="italic">"${rec.metadata.teacher_note}"</td>
          </tr>`
                )
                .join('')
        }
      </tbody>
    </table>

    <!-- Nhận xét của giáo viên -->
    <div style="margin-top: 15px; border: 1px dashed #475569; padding: 10px; background-color: #fafafa;">
      <p class="bold" style="margin: 0 0 5px 0;">III. NHẬN XÉT VÀ KHUYẾN NGHỊ CỦA GIÁO VIÊN BỘ MÔN:</p>
      <p style="margin: 0; line-height: 1.6;">
        - Học sinh <strong>${studentName}</strong> có tinh thần học tập tích cực, chủ động tư duy theo luồng Socratic.<br/>
        - Khả năng tự phát hiện lỗi sai trong Scratch tốt, phản hồi nhanh nhạy.<br/>
        - Khuyến nghị: Tiếp tục rèn luyện gõ bàn phím 10 ngón và khám phá các khối lệnh cảm biến nâng cao.
      </p>
    </div>

    <!-- Chữ ký -->
    <table class="footer-table">
      <tr>
        <td></td>
        <td>
          <em>Ngày ..... tháng ..... năm 2026</em><br/>
          <strong style="text-transform: uppercase;">Giáo viên bộ môn Tin học</strong><br/>
          <span style="font-size: 10pt; color: #64748b;">(Ký và ghi rõ họ tên)</span><br/><br/><br/><br/>
          <strong>........................................................</strong>
        </td>
      </tr>
    </table>
  </div>
</body>
</html>`;

  const blob = new Blob(['\uFEFF' + htmlContent], {
    type: 'application/msword;charset=utf-8;',
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  const safeName = studentName.replace(/\s+/g, '_');
  link.setAttribute('download', `Phieu_Nhan_Xet_Tin_Hoc_${studentId}_${safeName}.doc`);
  document.body.appendChild(link);
  link.click();
  link.remove();
}

/**
 * Xuất Báo cáo Tổng hợp Toàn bộ Học sinh trong Lớp ra Microsoft Word (.doc)
 */
export function exportClassSummaryDocx(
  students: Student[],
  records: AssessmentRecord[],
  grade: GradeLevel
): void {
  const htmlContent = `
<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head>
  <meta charset='utf-8'>
  <title>Báo Cáo Tổng Hợp Tin Học Lớp ${grade}</title>
  <style>
    @page WordSection1 { size: 21.0cm 29.7cm; margin: 2cm; }
    body { font-family: 'Times New Roman', serif; font-size: 12pt; line-height: 1.3; }
    h2 { text-align: center; text-transform: uppercase; color: #1e3a8a; }
    table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 11pt; }
    th, td { border: 1px solid #000; padding: 5px; text-align: left; }
    th { background-color: #f1f5f9; text-align: center; }
  </style>
</head>
<body>
  <h2>BÁO CÁO ĐÁNH GIÁ NĂNG LỰC HỌC SINH MÔN TIN HỌC - ${grade.toUpperCase()}</h2>
  <p style="text-align: center; font-style: italic;">Tổng số học sinh kết nối Google Sheet: ${students.length} em | Ngày xuất: ${new Date().toLocaleDateString('vi-VN')}</p>
  <table>
    <thead>
      <tr>
        <th style="width: 8%;">STT</th>
        <th style="width: 15%;">Mã HS</th>
        <th style="width: 25%;">Họ và tên</th>
        <th style="width: 12%;">Lớp</th>
        <th style="width: 12%;">Số sao ⭐</th>
        <th style="width: 28%;">Ghi chú chẩn đoán NUTA</th>
      </tr>
    </thead>
    <tbody>
      ${students
        .map(
          (st, idx) => `
        <tr>
          <td style="text-align: center;">${idx + 1}</td>
          <td style="text-align: center; font-family: monospace; font-weight: bold;">${st.id}</td>
          <td><strong>${st.name}</strong></td>
          <td style="text-align: center;">${st.className || st.grade}</td>
          <td style="text-align: center; font-weight: bold; color: #b45309;">${st.stars || 0}</td>
          <td style="font-style: italic;">${st.notes || 'Chăm chỉ hoàn thành bài tập'}</td>
        </tr>`
        )
        .join('')}
    </tbody>
  </table>
</body>
</html>`;

  const blob = new Blob(['\uFEFF' + htmlContent], {
    type: 'application/msword;charset=utf-8;',
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', `Bao_Cao_Tong_Hop_Tin_Hoc_${grade.replace(/\s+/g, '_')}_${Date.now()}.doc`);
  document.body.appendChild(link);
  link.click();
  link.remove();
}
