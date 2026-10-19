-- Chạy trong Supabase SQL Editor.
-- Script tìm hồ sơ theo tên BÙI HẢI BẰNG và dừng nếu không có đúng một hồ sơ khớp.
-- Nếu hồ sơ mang tên khác, sửa chuỗi trong điều kiện lower(trim(full_name)) bên dưới.

DO $cv$
DECLARE
  profile_ids uuid[];
  target_uid uuid;
  regina_id public.work_experiences.id%TYPE;
  sicix_id public.work_experiences.id%TYPE;
  next_order integer;
BEGIN
  SELECT array_agg(user_id)
    INTO profile_ids
    FROM public.cv_profiles
   WHERE lower(trim(full_name)) = lower('BÙI HẢI BẰNG');

  IF COALESCE(array_length(profile_ids, 1), 0) <> 1 THEN
    RAISE EXCEPTION 'Cần đúng một hồ sơ có tên BÙI HẢI BẰNG; tìm thấy % hồ sơ.',
      COALESCE(array_length(profile_ids, 1), 0);
  END IF;

  target_uid := profile_ids[1];

  SELECT id
    INTO regina_id
    FROM public.work_experiences
   WHERE user_id = target_uid
     AND role = 'Nhân viên hạng mục - Nhân viên phát triển hệ thống'
     AND company = 'TNHH Regina Miracle International Việt Nam'
     AND start_date = DATE '2024-08-20'
   ORDER BY sort_order
   LIMIT 1;

  SELECT id
    INTO sicix_id
    FROM public.work_experiences
   WHERE user_id = target_uid
     AND role = 'Thực tập sinh'
     AND company = 'Công ty cổ phần công nghệ Sicix'
     AND start_date = DATE '2023-08-04'
   ORDER BY sort_order
   LIMIT 1;

  IF regina_id IS NULL THEN
    SELECT COALESCE(MAX(sort_order), -1) + 1
      INTO next_order
      FROM public.work_experiences
     WHERE user_id = target_uid;

    INSERT INTO public.work_experiences
      (user_id, role, company, industry, location, start_date, end_date, is_current, sort_order)
    VALUES
      (target_uid,
       'Nhân viên hạng mục - Nhân viên phát triển hệ thống',
       'TNHH Regina Miracle International Việt Nam',
       NULL, NULL,
       DATE '2024-08-20', DATE '2026-02-27', false, next_order)
    RETURNING id INTO regina_id;
  END IF;

  IF NOT EXISTS (
    SELECT 1 FROM public.experience_bullets WHERE experience_id = regina_id
  ) THEN
    INSERT INTO public.experience_bullets (experience_id, user_id, content, sort_order)
    VALUES
      (regina_id, target_uid, 'Phân tích nghiệp vụ và vận hành hệ thống: Tiếp nhận và xử lý yêu cầu của các phòng ban qua hệ thống quản lý OA.', 0),
      (regina_id, target_uid, 'Hỗ trợ người dùng các hệ thống nội bộ: in tem IMMQI, xếp chuyền FR, SAP ERP, hệ thống sản xuất và báo biểu đơn hàng.', 1),
      (regina_id, target_uid, 'Phân tích yêu cầu nghiệp vụ, phối hợp với lập trình viên để triển khai giải pháp kỹ thuật.', 2),
      (regina_id, target_uid, 'Làm việc với nhà cung cấp phần mềm để bảo trì, nâng cấp hệ thống và xử lý sự cố.', 3),
      (regina_id, target_uid, 'Số hóa quy trình và báo cáo dữ liệu: Tự động hóa các quy trình xử lý dữ liệu trước đây thực hiện thủ công trên Excel.', 4),
      (regina_id, target_uid, 'Xây dựng các module tổng hợp báo cáo sản lượng, đơn hàng và dữ liệu sản xuất.', 5),
      (regina_id, target_uid, 'Tích hợp và xử lý dữ liệu từ hệ thống SAP theo nghiệp vụ thực tế.', 6),
      (regina_id, target_uid, 'Phát triển hệ thống nội bộ phục vụ quản lý sản xuất bằng C#/.NET.', 7);
  END IF;

  IF sicix_id IS NULL THEN
    SELECT COALESCE(MAX(sort_order), -1) + 1
      INTO next_order
      FROM public.work_experiences
     WHERE user_id = target_uid;

    INSERT INTO public.work_experiences
      (user_id, role, company, industry, location, start_date, end_date, is_current, sort_order)
    VALUES
      (target_uid,
       'Thực tập sinh',
       'Công ty cổ phần công nghệ Sicix',
       NULL, NULL,
       DATE '2023-08-04', DATE '2024-07-15', false, next_order)
    RETURNING id INTO sicix_id;
  END IF;

  IF NOT EXISTS (
    SELECT 1 FROM public.experience_bullets WHERE experience_id = sicix_id
  ) THEN
    INSERT INTO public.experience_bullets (experience_id, user_id, content, sort_order)
    VALUES
      (sicix_id, target_uid, 'Tham gia phân tích yêu cầu nghiệp vụ và làm rõ luồng xử lý hệ thống.', 0),
      (sicix_id, target_uid, 'Phát triển giao diện web và mobile theo yêu cầu dự án; tích hợp RESTful API, xử lý dữ liệu trên ứng dụng mobile.', 1),
      (sicix_id, target_uid, 'Thực hiện kiểm thử phần mềm, xây dựng test case, báo cáo lỗi.', 2),
      (sicix_id, target_uid, 'Hỗ trợ sửa lỗi và cải tiến tính năng trong quá trình phát triển.', 3);
  END IF;
END;
$cv$;
