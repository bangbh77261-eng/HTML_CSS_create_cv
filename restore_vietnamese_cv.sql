-- Restore the original Vietnamese CV from the supplied CSV.
-- Run this together with update_bui_hai_bang_cv_english.sql to keep both versions.
-- This script targets only the Vietnamese profile; it does not change the separate English profile.

BEGIN;

DO $restore$
DECLARE
  vn_uid uuid;
  profile_count integer;
  row_count integer;
  regina_id public.work_experiences.id%TYPE;
  sicix_id public.work_experiences.id%TYPE;
  enspire_id public.work_experiences.id%TYPE;
  group_id public.skill_groups.id%TYPE;
  grp record;
  skill_name text;
  group_index integer := 0;
  skill_index integer;
BEGIN
  SELECT count(*)
    INTO profile_count
    FROM public.cv_profiles
   WHERE lower(trim(full_name)) IN (lower('BÙI HẢI BẰNG'), lower('BUI HAI BANG'))
     AND role_tag IN (
       'Chuyên viên IT Phân tích nghiệp vụ & Cải tiến quy trình | ERP/SAP',
       'IT Support | Business Analysis | ERP / SAP'
     );

  IF profile_count <> 1 THEN
    RAISE EXCEPTION 'Expected one Vietnamese Bui Hai Bang profile, found %.', profile_count;
  END IF;

  SELECT user_id
    INTO vn_uid
    FROM public.cv_profiles
   WHERE lower(trim(full_name)) IN (lower('BÙI HẢI BẰNG'), lower('BUI HAI BANG'))
     AND role_tag IN (
       'Chuyên viên IT Phân tích nghiệp vụ & Cải tiến quy trình | ERP/SAP',
       'IT Support | Business Analysis | ERP / SAP'
     );

  UPDATE public.cv_profiles
     SET full_name = 'BÙI HẢI BẰNG',
         role_tag = 'Chuyên viên IT Phân tích nghiệp vụ & Cải tiến quy trình | ERP/SAP',
         summary = $vn_summary$Chuyên viên có gần 2 năm kinh nghiệm (bao gồm thực tập) trong môi trường doanh nghiệp sản xuất, tập trung vào phân tích yêu cầu nghiệp vụ, số hóa và cải tiến quy trình, báo cáo dữ liệu sản xuất và hỗ trợ vận hành hệ thống ERP/SAP. Quen phối hợp với nhiều phòng ban và nhà cung cấp phần mềm để thay thế các thao tác thủ công trên Excel bằng quy trình, hệ thống chuẩn hóa. Nền tảng kỹ sư công nghệ thông tin giúp trao đổi hiệu quả với đội kỹ thuật và nhanh chóng làm quen với hệ thống mới.
Mong muốn đảm nhận các vị trí như phân tích nghiệp vụ, điều phối dự án, vận hành/hỗ trợ ERP, phân tích dữ liệu - báo cáo hoặc cải tiến quy trình trong các lĩnh vực sản xuất, logistics, thương mại và kỹ thuật.$vn_summary$,
         monogram = 'HB',
         layout = jsonb_set(
           COALESCE(layout, '{}'::jsonb),
           '{sections}',
           COALESCE(layout->'sections', '{}'::jsonb) || jsonb_build_object(
             'titles',
             COALESCE(layout #> '{sections,titles}', '{}'::jsonb) || jsonb_build_object(
               'contact', 'Liên hệ',
               'skills', 'Kỹ năng',
               'educations', 'Học vấn',
               'certifications', 'Chứng chỉ',
               'languages', 'Ngôn ngữ',
               'summary', 'Giới thiệu',
               'experiences', 'Kinh nghiệm làm việc',
               'projects', 'Dự án nổi bật'
             )
           ),
           true
         )
   WHERE user_id = vn_uid;

  -- Restore the three work experience records and their dates.
  SELECT count(*) INTO row_count
    FROM public.work_experiences
   WHERE user_id = vn_uid AND start_date = DATE '2024-08-20';
  IF row_count > 1 THEN
    RAISE EXCEPTION 'More than one Regina Miracle experience has start date 2024-08-20.';
  ELSIF row_count = 0 THEN
    INSERT INTO public.work_experiences
      (user_id, role, company, industry, location, start_date, end_date, is_current, sort_order)
    VALUES
      (vn_uid, 'Nhân viên hạng mục - Nhân viên phát triển hệ thống',
       'TNHH Regina Miracle International Việt Nam', 'Nhà máy B', 'Hải Phòng',
       DATE '2024-08-20', DATE '2026-02-27', false, 1)
    RETURNING id INTO regina_id;
  ELSE
    SELECT id INTO regina_id FROM public.work_experiences
     WHERE user_id = vn_uid AND start_date = DATE '2024-08-20';
    UPDATE public.work_experiences
       SET role = 'Nhân viên hạng mục - Nhân viên phát triển hệ thống',
           company = 'TNHH Regina Miracle International Việt Nam',
           industry = 'Nhà máy B', location = 'Hải Phòng',
           end_date = DATE '2026-02-27', is_current = false, sort_order = 1
     WHERE id = regina_id;
  END IF;

  DELETE FROM public.experience_bullets
   WHERE user_id = vn_uid AND experience_id = regina_id;
  INSERT INTO public.experience_bullets (experience_id, user_id, content, sort_order)
  VALUES
    (regina_id, vn_uid, 'Phân tích nghiệp vụ và vận hành hệ thống: Tiếp nhận và xử lý yêu cầu của các phòng ban qua hệ thống quản lý OA.', 0),
    (regina_id, vn_uid, 'Hỗ trợ người dùng các hệ thống nội bộ: in tem IMMQI, xếp chuyền FR, SAP ERP, hệ thống sản xuất và báo biểu đơn hàng.', 1),
    (regina_id, vn_uid, 'Phân tích yêu cầu nghiệp vụ, phối hợp với lập trình viên để triển khai giải pháp kỹ thuật.', 2),
    (regina_id, vn_uid, 'Làm việc với nhà cung cấp phần mềm để bảo trì, nâng cấp hệ thống và xử lý sự cố.', 3),
    (regina_id, vn_uid, 'Số hóa quy trình và báo cáo dữ liệu: Tự động hóa các quy trình xử lý dữ liệu trước đây thực hiện thủ công trên Excel.', 4),
    (regina_id, vn_uid, 'Xây dựng các module tổng hợp báo cáo sản lượng, đơn hàng và dữ liệu sản xuất.', 5),
    (regina_id, vn_uid, 'Tích hợp và xử lý dữ liệu từ hệ thống SAP theo nghiệp vụ thực tế.', 6),
    (regina_id, vn_uid, 'Phát triển hệ thống nội bộ phục vụ quản lý sản xuất bằng C#/.NET.', 7);

  SELECT count(*) INTO row_count
    FROM public.work_experiences
   WHERE user_id = vn_uid AND start_date = DATE '2023-08-04';
  IF row_count > 1 THEN
    RAISE EXCEPTION 'More than one Sicix internship has start date 2023-08-04.';
  ELSIF row_count = 0 THEN
    INSERT INTO public.work_experiences
      (user_id, role, company, industry, location, start_date, end_date, is_current, sort_order)
    VALUES
      (vn_uid, 'Thực tập sinh', 'Công ty cổ phần công nghệ Sicix',
       'Chuyển đổi số', 'Hải Phòng', DATE '2023-08-04', DATE '2024-07-15', false, 2)
    RETURNING id INTO sicix_id;
  ELSE
    SELECT id INTO sicix_id FROM public.work_experiences
     WHERE user_id = vn_uid AND start_date = DATE '2023-08-04';
    UPDATE public.work_experiences
       SET role = 'Thực tập sinh', company = 'Công ty cổ phần công nghệ Sicix',
           industry = 'Chuyển đổi số', location = 'Hải Phòng',
           end_date = DATE '2024-07-15', is_current = false, sort_order = 2
     WHERE id = sicix_id;
  END IF;

  DELETE FROM public.experience_bullets
   WHERE user_id = vn_uid AND experience_id = sicix_id;
  INSERT INTO public.experience_bullets (experience_id, user_id, content, sort_order)
  VALUES
    (sicix_id, vn_uid, 'Tham gia phân tích yêu cầu nghiệp vụ và làm rõ luồng xử lý hệ thống.', 0),
    (sicix_id, vn_uid, 'Phát triển giao diện web và mobile theo yêu cầu dự án; tích hợp RESTful API, xử lý dữ liệu trên ứng dụng mobile.', 1),
    (sicix_id, vn_uid, 'Thực hiện kiểm thử phần mềm, xây dựng test case, báo cáo lỗi.', 2),
    (sicix_id, vn_uid, 'Hỗ trợ sửa lỗi và cải tiến tính năng trong quá trình phát triển.', 3);

  SELECT count(*) INTO row_count
    FROM public.work_experiences
   WHERE user_id = vn_uid AND start_date = DATE '2026-05-25';
  IF row_count > 1 THEN
    RAISE EXCEPTION 'More than one Enspire experience has start date 2026-05-25.';
  ELSIF row_count = 0 THEN
    INSERT INTO public.work_experiences
      (user_id, role, company, industry, location, start_date, end_date, is_current, sort_order)
    VALUES
      (vn_uid, 'Nhân viên IT', 'Trung tâm ngoại ngữ Enspire',
       'Giáo dục tiếng anh', 'Hải Phòng', DATE '2026-05-25', NULL, true, 0)
    RETURNING id INTO enspire_id;
  ELSE
    SELECT id INTO enspire_id FROM public.work_experiences
     WHERE user_id = vn_uid AND start_date = DATE '2026-05-25';
    UPDATE public.work_experiences
       SET role = 'Nhân viên IT', company = 'Trung tâm ngoại ngữ Enspire',
           industry = 'Giáo dục tiếng anh', location = 'Hải Phòng',
           end_date = NULL, is_current = true, sort_order = 0
     WHERE id = enspire_id;
  END IF;

  DELETE FROM public.experience_bullets
   WHERE user_id = vn_uid AND experience_id = enspire_id;
  INSERT INTO public.experience_bullets (experience_id, user_id, content, sort_order)
  VALUES
    (enspire_id, vn_uid, 'Kiểm tra, khắc phục các sự cố mạng LAN, Wi-Fi, kết nối Internet và thiết bị ngoại vi.', 0),
    (enspire_id, vn_uid, 'Cài đặt, cấu hình và xử lý sự cố máy tính, máy in, máy chiếu, thiết bị âm thanh và các thiết bị phục vụ giảng dạy.', 1),
    (enspire_id, vn_uid, 'Hỗ trợ kỹ thuật cho giáo viên, nhân viên và học viên trong quá trình sử dụng máy tính, phần mềm và các thiết bị công nghệ.', 2),
    (enspire_id, vn_uid, 'Hỗ trợ sử dụng, vận hành và xử lý lỗi cơ bản trên các phần mềm quản lý đào tạo, quản lý học viên và hệ thống học trực tuyến.', 3);

  -- Restore the original project names and descriptions. Expolander was part of the Sicix project in the CSV.
  UPDATE public.projects
     SET title = 'Hệ thống quản lý in tem sản xuất - Bộ phận Hot Melt, Regina Miracle',
         role_meta = '06/2025 - 11/2025', year = 2025,
         description = $project_one$Vai trò: Lập trình viên / Phân tích nghiệp vụ. Hệ thống quản lý in tem cho các lệnh sản xuất trong nhà máy, thay thế quy trình thao tác thủ công trên Excel.
• Thiết kế và phát triển hệ thống bằng C# WinForms và WPF.
• Chức năng chính: quản lý người dùng và phân quyền; quản lý sản lượng theo lệnh sản xuất; thống kê sản lượng theo ngày; lưu lịch sử in tem; báo cáo dữ liệu sản xuất.
• Tích hợp dữ liệu từ hệ thống SAP.$project_one$,
         result = NULL, sort_order = 0
   WHERE user_id = vn_uid
     AND title IN ('Hệ thống quản lý in tem sản xuất - Bộ phận Hot Melt, Regina Miracle', 'Production Label Printing System');
  IF NOT FOUND THEN
    INSERT INTO public.projects (user_id, title, role_meta, year, description, result, sort_order)
    VALUES (vn_uid, 'Hệ thống quản lý in tem sản xuất - Bộ phận Hot Melt, Regina Miracle', '06/2025 - 11/2025', 2025,
            $project_one$Vai trò: Lập trình viên / Phân tích nghiệp vụ. Hệ thống quản lý in tem cho các lệnh sản xuất trong nhà máy, thay thế quy trình thao tác thủ công trên Excel.
• Thiết kế và phát triển hệ thống bằng C# WinForms và WPF.
• Chức năng chính: quản lý người dùng và phân quyền; quản lý sản lượng theo lệnh sản xuất; thống kê sản lượng theo ngày; lưu lịch sử in tem; báo cáo dữ liệu sản xuất.
• Tích hợp dữ liệu từ hệ thống SAP.$project_one$, NULL, 0);
  END IF;

  UPDATE public.projects
     SET title = 'Hệ thống gửi thông báo email tự động - Bộ phận IT, Regina Miracle',
         role_meta = '08/2025 - 12/2025', year = 2025,
         description = $project_two$Vai trò: Lập trình viên / Phân tích nghiệp vụ.
• Xây dựng chương trình tự động hóa (RPA) bằng C#/.NET để gửi email thông báo cho người dùng trong doanh nghiệp.
• Kiểm tra, đối soát dữ liệu từ hệ thống nhân sự; xác định danh sách nhân viên nghỉ việc và đối tượng cần nhận thông báo.
• Cấu hình lịch chạy tự động và xử lý logic nghiệp vụ.$project_two$,
         result = NULL, sort_order = 1
   WHERE user_id = vn_uid
     AND title IN ('Hệ thống gửi thông báo email tự động - Bộ phận IT, Regina Miracle', 'Automated Email Notice System');
  IF NOT FOUND THEN
    INSERT INTO public.projects (user_id, title, role_meta, year, description, result, sort_order)
    VALUES (vn_uid, 'Hệ thống gửi thông báo email tự động - Bộ phận IT, Regina Miracle', '08/2025 - 12/2025', 2025,
            $project_two$Vai trò: Lập trình viên / Phân tích nghiệp vụ.
• Xây dựng chương trình tự động hóa (RPA) bằng C#/.NET để gửi email thông báo cho người dùng trong doanh nghiệp.
• Kiểm tra, đối soát dữ liệu từ hệ thống nhân sự; xác định danh sách nhân viên nghỉ việc và đối tượng cần nhận thông báo.
• Cấu hình lịch chạy tự động và xử lý logic nghiệp vụ.$project_two$, NULL, 1);
  END IF;

  UPDATE public.projects
     SET title = 'Hệ thống quản lý tem xả vải - Regina Miracle',
         role_meta = '11/2025 - 12/2025', year = 2025,
         description = $project_three$Vai trò: Phân tích nghiệp vụ.
• Thu thập, phân tích dữ liệu từ hệ thống SAP và kiểm tra dữ liệu giữa các môi trường.
• Xây dựng luồng xử lý nghiệp vụ và tài liệu hệ thống.
• Theo dõi tiến độ và báo cáo tình trạng dự án.$project_three$,
         result = NULL, sort_order = 2
   WHERE user_id = vn_uid
     AND title IN ('Hệ thống quản lý tem xả vải - Regina Miracle', 'Fabric Softener Label System');
  IF NOT FOUND THEN
    INSERT INTO public.projects (user_id, title, role_meta, year, description, result, sort_order)
    VALUES (vn_uid, 'Hệ thống quản lý tem xả vải - Regina Miracle', '11/2025 - 12/2025', 2025,
            $project_three$Vai trò: Phân tích nghiệp vụ.
• Thu thập, phân tích dữ liệu từ hệ thống SAP và kiểm tra dữ liệu giữa các môi trường.
• Xây dựng luồng xử lý nghiệp vụ và tài liệu hệ thống.
• Theo dõi tiến độ và báo cáo tình trạng dự án.$project_three$, NULL, 2);
  END IF;

  UPDATE public.projects
     SET title = 'Sicix Mobile App - Công ty cổ phần công nghệ Sicix',
         role_meta = '11/2023 - 07/2024', year = 2024,
         description = $project_four$Vai trò: Developer / Tester. Ứng dụng mobile phục vụ quản lý nội bộ. Quy mô nhóm: 6 thành viên.
• Phát triển tính năng bằng Flutter, tích hợp API backend và xử lý dữ liệu.
• Thực hiện kiểm thử và sửa lỗi.
Ứng dụng quản lý sự kiện Expolander | 01/2024 - 07/2024
Vai trò: Mobile Developer. Ứng dụng cho phép người dùng quản lý sự kiện và lịch họp. Quy mô nhóm: 4 thành viên.
• Phát triển giao diện người dùng, sửa lỗi và cải tiến tính năng theo yêu cầu.$project_four$,
         result = NULL, sort_order = 3
   WHERE user_id = vn_uid
     AND title IN ('Sicix Mobile App - Công ty cổ phần công nghệ Sicix', 'Sicix Mobile App');
  IF NOT FOUND THEN
    INSERT INTO public.projects (user_id, title, role_meta, year, description, result, sort_order)
    VALUES (vn_uid, 'Sicix Mobile App - Công ty cổ phần công nghệ Sicix', '11/2023 - 07/2024', 2024,
            $project_four$Vai trò: Developer / Tester. Ứng dụng mobile phục vụ quản lý nội bộ. Quy mô nhóm: 6 thành viên.
• Phát triển tính năng bằng Flutter, tích hợp API backend và xử lý dữ liệu.
• Thực hiện kiểm thử và sửa lỗi.
Ứng dụng quản lý sự kiện Expolander | 01/2024 - 07/2024
Vai trò: Mobile Developer. Ứng dụng cho phép người dùng quản lý sự kiện và lịch họp. Quy mô nhóm: 4 thành viên.
• Phát triển giao diện người dùng, sửa lỗi và cải tiến tính năng theo yêu cầu.$project_four$, NULL, 3);
  END IF;

  DELETE FROM public.projects
   WHERE user_id = vn_uid AND title = 'Expolander Event Management App';

  -- Restore all skill groups and skills from the Vietnamese CSV.
  DELETE FROM public.skills WHERE user_id = vn_uid;
  DELETE FROM public.skill_groups WHERE user_id = vn_uid;

  FOR grp IN
    SELECT * FROM (VALUES
      ('PHÂN TÍCH NGHIỆP VỤ', ARRAY['Thu thập và phân tích yêu cầu, xây dựng luồng nghiệp vụ, viết tài liệu hệ thống, theo dõi tiến độ và báo cáo tình trạng dự án.']::text[]),
      ('SỐ HÓA VÀ CẢI TIẾN QUY TRÌNH', ARRAY['Chuyển quy trình thủ công trên Excel sang hệ thống phần mềm; tự động hóa bằng RPA/Automation và VBA.']::text[]),
      ('DỮ LIỆU VÀ BÁO CÁO', ARRAY['Tổng hợp, đối soát và báo cáo sản lượng, đơn hàng, dữ liệu sản xuất; làm việc với cơ sở dữ liệu SQL Server, MySQL, PostgreSQL.']::text[]),
      ('ERP VÀ HỆ THỐNG DOANH NGHIỆP', ARRAY['Tích hợp và xử lý dữ liệu từ SAP; hỗ trợ người dùng; làm việc với nhà cung cấp phần mềm.']::text[]),
      ('LÀM VIỆC NHÓM VÀ GIAO TIẾP', ARRAY['Tiếp nhận yêu cầu từ nhiều phòng ban, làm cầu nối giữa người dùng nghiệp vụ và đội kỹ thuật.']::text[]),
      ('CÔNG CỤ', ARRAY['Microsoft Office, VBA, Git, Postman, Figma; lập trình C#/.NET, Flutter, JavaScript, RESTful API.']::text[])
    ) AS skill_data(group_name, skill_names)
  LOOP
    INSERT INTO public.skill_groups (user_id, name, sort_order)
    VALUES (vn_uid, grp.group_name, group_index)
    RETURNING id INTO group_id;

    skill_index := 0;
    FOREACH skill_name IN ARRAY grp.skill_names LOOP
      INSERT INTO public.skills (group_id, user_id, name, sort_order)
      VALUES (group_id, vn_uid, skill_name, skill_index);
      skill_index := skill_index + 1;
    END LOOP;
    group_index := group_index + 1;
  END LOOP;

  UPDATE public.educations
     SET school = 'Trường Đại học Hàng hải Việt Nam',
         degree = 'Kỹ sư - Công nghệ phần mềm (CNTT)',
         notes = 'Tham gia các dự án thực tế từ sớm và làm việc trong môi trường doanh nghiệp chuyên nghiệp.'
   WHERE user_id = vn_uid
     AND school IN ('Trường Đại học Hàng hải Việt Nam', 'Vietnam Maritime University');
  IF NOT FOUND THEN
    INSERT INTO public.educations (user_id, school, degree, start_year, end_year, notes, sort_order)
    VALUES (vn_uid, 'Trường Đại học Hàng hải Việt Nam', 'Kỹ sư - Công nghệ phần mềm (CNTT)',
            NULL, NULL, 'Tham gia các dự án thực tế từ sớm và làm việc trong môi trường doanh nghiệp chuyên nghiệp.', 0);
  END IF;

  DELETE FROM public.languages
   WHERE user_id = vn_uid AND name IN ('Tiếng Trung', 'Chinese', 'Tiếng Anh', 'English');
  INSERT INTO public.languages (user_id, name, level, sort_order)
  VALUES
    (vn_uid, 'Tiếng Trung', 'HSK 3', 0),
    (vn_uid, 'Tiếng Anh', 'Đọc hiểu tài liệu', 1);
END;
$restore$;

COMMIT;
