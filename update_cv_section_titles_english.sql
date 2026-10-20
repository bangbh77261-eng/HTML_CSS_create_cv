-- Translate the section headings on the separate English Bui Hai Bang CV.
-- The Vietnamese CV keeps its original Vietnamese headings.

DO $cv$
DECLARE
  target_uid uuid;
  profile_count integer;
BEGIN
  SELECT count(*)
    INTO profile_count
    FROM public.cv_profiles
   WHERE lower(trim(full_name)) IN (lower('BÙI HẢI BẰNG'), lower('BUI HAI BANG'))
     AND role_tag IN (
       'Business Analyst | Process Improvement | ERP/SAP | Production Data Reporting',
       'IT Specialist | Business Analysis & Process Improvement | ERP/SAP'
     );

  IF profile_count <> 1 THEN
    RAISE EXCEPTION 'Expected one separate English Bui Hai Bang profile, found %.', profile_count;
  END IF;

  SELECT user_id
    INTO target_uid
    FROM public.cv_profiles
   WHERE lower(trim(full_name)) IN (lower('BÙI HẢI BẰNG'), lower('BUI HAI BANG'))
     AND role_tag IN (
       'Business Analyst | Process Improvement | ERP/SAP | Production Data Reporting',
       'IT Specialist | Business Analysis & Process Improvement | ERP/SAP'
     );

  UPDATE public.cv_profiles
     SET layout = jsonb_set(
       COALESCE(layout, '{}'::jsonb),
       '{sections}',
       COALESCE(layout->'sections', '{}'::jsonb) || jsonb_build_object(
         'titles',
         COALESCE(layout #> '{sections,titles}', '{}'::jsonb) || jsonb_build_object(
           'contact', 'Contact',
           'skills', 'Skills',
           'educations', 'Education',
           'certifications', 'Certificates',
           'languages', 'Languages',
           'summary', 'About Me',
           'experiences', 'Work Experience',
           'projects', 'Projects'
         )
       ),
       true
     )
   WHERE user_id = target_uid;
END;
$cv$;
