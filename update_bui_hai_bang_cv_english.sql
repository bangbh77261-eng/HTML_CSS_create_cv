-- Update the Vietnamese Bui Hai Bang CV to simple English.
-- Run in Supabase SQL Editor. The script checks that one matching profile exists.
-- It keeps contact and certification data unchanged because the CSV has no values for them.

BEGIN;

DO $cv$
DECLARE
  target_uid uuid;
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
     AND (role_tag ILIKE '%Chuyên viên IT%' OR role_tag = 'IT Support | Business Analysis | ERP / SAP');

  IF profile_count <> 1 THEN
    RAISE EXCEPTION 'Expected one Vietnamese Bui Hai Bang profile, found %.', profile_count;
  END IF;

  SELECT user_id
    INTO target_uid
    FROM public.cv_profiles
   WHERE lower(trim(full_name)) IN (lower('BÙI HẢI BẰNG'), lower('BUI HAI BANG'))
     AND (role_tag ILIKE '%Chuyên viên IT%' OR role_tag = 'IT Support | Business Analysis | ERP / SAP');

  UPDATE public.cv_profiles
     SET full_name = 'BUI HAI BANG',
         role_tag = 'IT Support | Business Analysis | ERP / SAP',
         summary = 'I have almost two years of experience, including an internship, in manufacturing. I study business needs, improve work steps, prepare reports, and support SAP and other business systems. I work with teams and software providers to move manual Excel tasks to clear system steps. I have an IT degree and can work with technical teams and learn new systems quickly. I am looking for jobs in business analysis, project support, ERP support, data reports, or process improvement.',
         monogram = 'HB'
   WHERE user_id = target_uid;

  -- Regina Miracle experience
  SELECT count(*) INTO row_count
    FROM public.work_experiences
   WHERE user_id = target_uid AND start_date = DATE '2024-08-20';
  IF row_count > 1 THEN
    RAISE EXCEPTION 'More than one Regina Miracle experience has start date 2024-08-20.';
  ELSIF row_count = 0 THEN
    INSERT INTO public.work_experiences
      (user_id, role, company, industry, location, start_date, end_date, is_current, sort_order)
    VALUES
      (target_uid, 'System Development Staff / Business Analyst',
       'Regina Miracle International Vietnam Co., Ltd.', 'Manufacturing', 'Hai Phong, Vietnam',
       DATE '2024-08-20', DATE '2026-02-27', false, 1)
    RETURNING id INTO regina_id;
  ELSE
    SELECT id INTO regina_id FROM public.work_experiences
     WHERE user_id = target_uid AND start_date = DATE '2024-08-20';
    UPDATE public.work_experiences
       SET role = 'System Development Staff / Business Analyst',
           company = 'Regina Miracle International Vietnam Co., Ltd.',
           industry = 'Manufacturing', location = 'Hai Phong, Vietnam',
           end_date = DATE '2026-02-27', is_current = false, sort_order = 1
     WHERE id = regina_id;
  END IF;

  DELETE FROM public.experience_bullets
   WHERE user_id = target_uid AND experience_id = regina_id;
  INSERT INTO public.experience_bullets (experience_id, user_id, content, sort_order)
  VALUES
    (regina_id, target_uid, 'Received and handled requests from different departments through the OA system.', 0),
    (regina_id, target_uid, 'Helped users with IMMQI label printing, the FR line system, SAP ERP, production systems, and order reports.', 1),
    (regina_id, target_uid, 'Studied business needs and worked with developers to build technical solutions.', 2),
    (regina_id, target_uid, 'Worked with software vendors to maintain systems, improve features, and fix problems.', 3),
    (regina_id, target_uid, 'Moved manual Excel work to software and made data tasks faster.', 4),
    (regina_id, target_uid, 'Built reports for production output, orders, and factory data.', 5),
    (regina_id, target_uid, 'Used SAP data to support daily business work.', 6),
    (regina_id, target_uid, 'Built an internal production system with C#/.NET.', 7);

  -- Sicix internship
  SELECT count(*) INTO row_count
    FROM public.work_experiences
   WHERE user_id = target_uid AND start_date = DATE '2023-08-04';
  IF row_count > 1 THEN
    RAISE EXCEPTION 'More than one Sicix internship has start date 2023-08-04.';
  ELSIF row_count = 0 THEN
    INSERT INTO public.work_experiences
      (user_id, role, company, industry, location, start_date, end_date, is_current, sort_order)
    VALUES
      (target_uid, 'Software Development Intern', 'Sicix Technology Joint Stock Company',
       'Software Development', 'Hai Phong, Vietnam', DATE '2023-08-04', DATE '2024-07-15', false, 2)
    RETURNING id INTO sicix_id;
  ELSE
    SELECT id INTO sicix_id FROM public.work_experiences
     WHERE user_id = target_uid AND start_date = DATE '2023-08-04';
    UPDATE public.work_experiences
       SET role = 'Software Development Intern',
           company = 'Sicix Technology Joint Stock Company',
           industry = 'Software Development', location = 'Hai Phong, Vietnam',
           end_date = DATE '2024-07-15', is_current = false, sort_order = 2
     WHERE id = sicix_id;
  END IF;

  DELETE FROM public.experience_bullets
   WHERE user_id = target_uid AND experience_id = sicix_id;
  INSERT INTO public.experience_bullets (experience_id, user_id, content, sort_order)
  VALUES
    (sicix_id, target_uid, 'Helped study business needs and explain how the system should work.', 0),
    (sicix_id, target_uid, 'Built web and mobile screens, connected REST APIs, and worked with app data.', 1),
    (sicix_id, target_uid, 'Tested software, wrote test cases, and reported bugs.', 2),
    (sicix_id, target_uid, 'Fixed bugs and improved features with the team.', 3);

  -- Current IT Support role. The source CSV had one repeated bullet; it is included once here.
  SELECT count(*) INTO row_count
    FROM public.work_experiences
   WHERE user_id = target_uid AND start_date = DATE '2026-05-25';
  IF row_count > 1 THEN
    RAISE EXCEPTION 'More than one Enspire experience has start date 2026-05-25.';
  ELSIF row_count = 0 THEN
    INSERT INTO public.work_experiences
      (user_id, role, company, industry, location, start_date, end_date, is_current, sort_order)
    VALUES
      (target_uid, 'IT Support Staff', 'Enspire English Center',
       'Education', 'Hai Phong, Vietnam', DATE '2026-05-25', NULL, true, 0)
    RETURNING id INTO enspire_id;
  ELSE
    SELECT id INTO enspire_id FROM public.work_experiences
     WHERE user_id = target_uid AND start_date = DATE '2026-05-25';
    UPDATE public.work_experiences
       SET role = 'IT Support Staff', company = 'Enspire English Center',
           industry = 'Education', location = 'Hai Phong, Vietnam',
           end_date = NULL, is_current = true, sort_order = 0
     WHERE id = enspire_id;
  END IF;

  DELETE FROM public.experience_bullets
   WHERE user_id = target_uid AND experience_id = enspire_id;
  INSERT INTO public.experience_bullets (experience_id, user_id, content, sort_order)
  VALUES
    (enspire_id, target_uid, 'Checked and fixed LAN, Wi-Fi, internet, and device connection problems.', 0),
    (enspire_id, target_uid, 'Set up and fixed computers, printers, projectors, sound systems, and teaching devices.', 1),
    (enspire_id, target_uid, 'Helped teachers, staff, and students use computers, software, and technology devices.', 2),
    (enspire_id, target_uid, 'Helped users with training, student, and online learning systems and fixed basic problems.', 3);

  -- Projects: update the four projects from the CSV and add Expolander as its own project.
  UPDATE public.projects
     SET title = 'Production Label Printing System',
         role_meta = 'Developer / Business Analyst | Jun 2025 - Nov 2025',
         year = 2025,
         description = 'This system prints labels for factory orders and replaces manual Excel work. Built with C# WinForms and WPF. Main features: user access, output by production order, daily reports, print history, and SAP data.',
         result = NULL, sort_order = 2
   WHERE user_id = target_uid
     AND title IN ('Hệ thống quản lý in tem sản xuất - Bộ phận Hot Melt, Regina Miracle', 'Production Label Printing System');
  IF NOT FOUND THEN
    INSERT INTO public.projects (user_id, title, role_meta, year, description, result, sort_order)
    VALUES (target_uid, 'Production Label Printing System', 'Developer / Business Analyst | Jun 2025 - Nov 2025', 2025,
            'This system prints labels for factory orders and replaces manual Excel work. Built with C# WinForms and WPF. Main features: user access, output by production order, daily reports, print history, and SAP data.', NULL, 2);
  END IF;

  UPDATE public.projects
     SET title = 'Automated Email Notice System',
         role_meta = 'Developer / Business Analyst | Aug 2025 - Dec 2025',
         year = 2025,
         description = 'Built a C#/.NET tool to send email notices to staff. Checked HR data to find staff who left and the right people to receive each notice. Set the run time and work rules.',
         result = NULL, sort_order = 0
   WHERE user_id = target_uid
     AND title IN ('Hệ thống gửi thông báo email tự động - Bộ phận IT, Regina Miracle', 'Automated Email Notice System');
  IF NOT FOUND THEN
    INSERT INTO public.projects (user_id, title, role_meta, year, description, result, sort_order)
    VALUES (target_uid, 'Automated Email Notice System', 'Developer / Business Analyst | Aug 2025 - Dec 2025', 2025,
            'Built a C#/.NET tool to send email notices to staff. Checked HR data to find staff who left and the right people to receive each notice. Set the run time and work rules.', NULL, 0);
  END IF;

  UPDATE public.projects
     SET title = 'Fabric Softener Label System',
         role_meta = 'Business Analyst | Nov 2025 - Dec 2025',
         year = 2025,
         description = 'Checked SAP data and compared data from different systems. Planned the work steps and wrote system documents. Followed project work and shared progress with the team.',
         result = NULL, sort_order = 1
   WHERE user_id = target_uid
     AND title IN ('Hệ thống quản lý tem xả vải - Regina Miracle', 'Fabric Softener Label System');
  IF NOT FOUND THEN
    INSERT INTO public.projects (user_id, title, role_meta, year, description, result, sort_order)
    VALUES (target_uid, 'Fabric Softener Label System', 'Business Analyst | Nov 2025 - Dec 2025', 2025,
            'Checked SAP data and compared data from different systems. Planned the work steps and wrote system documents. Followed project work and shared progress with the team.', NULL, 1);
  END IF;

  UPDATE public.projects
     SET title = 'Sicix Mobile App',
         role_meta = 'Mobile Developer / Tester | Nov 2023 - Jul 2024',
         year = 2024,
         description = 'Team: 6 people. Built mobile features with Flutter, connected the app to the server API, tested features, and fixed bugs.',
         result = NULL, sort_order = 3
   WHERE user_id = target_uid
     AND title IN ('Sicix Mobile App - Công ty cổ phần công nghệ Sicix', 'Sicix Mobile App');
  IF NOT FOUND THEN
    INSERT INTO public.projects (user_id, title, role_meta, year, description, result, sort_order)
    VALUES (target_uid, 'Sicix Mobile App', 'Mobile Developer / Tester | Nov 2023 - Jul 2024', 2024,
            'Team: 6 people. Built mobile features with Flutter, connected the app to the server API, tested features, and fixed bugs.', NULL, 3);
  END IF;

  IF NOT EXISTS (
    SELECT 1 FROM public.projects
     WHERE user_id = target_uid AND title = 'Expolander Event Management App'
  ) THEN
    INSERT INTO public.projects (user_id, title, role_meta, year, description, result, sort_order)
    VALUES (target_uid, 'Expolander Event Management App', 'Mobile Developer | Jan 2024 - Jul 2024', 2024,
            'Team: 4 people. Built app screens for event and meeting plans. Fixed bugs and improved features based on project needs.', NULL, 4);
  ELSE
    UPDATE public.projects
       SET role_meta = 'Mobile Developer | Jan 2024 - Jul 2024', year = 2024,
           description = 'Team: 4 people. Built app screens for event and meeting plans. Fixed bugs and improved features based on project needs.',
           result = NULL, sort_order = 4
     WHERE user_id = target_uid AND title = 'Expolander Event Management App';
  END IF;

  -- Replace skill groups and skills with the English list from the CV.
  DELETE FROM public.skills WHERE user_id = target_uid;
  DELETE FROM public.skill_groups WHERE user_id = target_uid;

  FOR grp IN
    SELECT * FROM (VALUES
      ('Business Analysis', ARRAY['Collect user needs', 'Write system documents', 'Track work and report progress']::text[]),
      ('Process Improvement', ARRAY['Move Excel work to software', 'Automate tasks with RPA and VBA']::text[]),
      ('Data and Reports', ARRAY['Check production and order data', 'Prepare reports', 'SQL Server, MySQL, PostgreSQL']::text[]),
      ('ERP and Business Systems', ARRAY['SAP data support', 'Help system users', 'Work with software vendors']::text[]),
      ('Tools and Programming', ARRAY['Microsoft Office', 'C#/.NET', 'Flutter', 'JavaScript', 'REST API', 'Git', 'Postman', 'Figma']::text[])
    ) AS skill_data(group_name, skill_names)
  LOOP
    INSERT INTO public.skill_groups (user_id, name, sort_order)
    VALUES (target_uid, grp.group_name, group_index)
    RETURNING id INTO group_id;

    skill_index := 0;
    FOREACH skill_name IN ARRAY grp.skill_names LOOP
      INSERT INTO public.skills (group_id, user_id, name, sort_order)
      VALUES (group_id, target_uid, skill_name, skill_index);
      skill_index := skill_index + 1;
    END LOOP;
    group_index := group_index + 1;
  END LOOP;

  -- Translate the existing education row, or add it if it is missing.
  UPDATE public.educations
     SET school = 'Vietnam Maritime University',
         degree = 'Bachelor of Engineering in Software Engineering',
         notes = 'Joined real projects early and learned to work in a professional team.'
   WHERE user_id = target_uid
     AND school IN ('Trường Đại học Hàng hải Việt Nam', 'Vietnam Maritime University');
  IF NOT FOUND THEN
    INSERT INTO public.educations (user_id, school, degree, start_year, end_year, notes, sort_order)
    VALUES (target_uid, 'Vietnam Maritime University', 'Bachelor of Engineering in Software Engineering',
            NULL, NULL, 'Joined real projects early and learned to work in a professional team.', 0);
  END IF;

  -- Keep the actual language level; do not claim a TOEIC score that is not in the source data.
  UPDATE public.languages
     SET name = 'Chinese', level = 'HSK 3', sort_order = 0
   WHERE user_id = target_uid AND name IN ('Tiếng Trung', 'Chinese');
  IF NOT FOUND THEN
    INSERT INTO public.languages (user_id, name, level, sort_order)
    VALUES (target_uid, 'Chinese', 'HSK 3', 0);
  END IF;

  UPDATE public.languages
     SET name = 'English', level = 'Basic English reading; can read simple technical documents', sort_order = 1
   WHERE user_id = target_uid AND name IN ('Tiếng Anh', 'English');
  IF NOT FOUND THEN
    INSERT INTO public.languages (user_id, name, level, sort_order)
    VALUES (target_uid, 'English', 'Basic English reading; can read simple technical documents', 1);
  END IF;
END;
$cv$;

COMMIT;
