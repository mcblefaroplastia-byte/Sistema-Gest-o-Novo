-- WARNING: This schema is for context only and is not meant to be run.
-- Table order and constraints may not be valid for execution.

CREATE TABLE public.profiles (
  id uuid NOT NULL,
  full_name text NOT NULL,
  role USER-DEFINED NOT NULL DEFAULT 'orientadora'::app_role,
  active boolean NOT NULL DEFAULT true,
  created_at timestamp with time zone NOT NULL DEFAULT now(),
  CONSTRAINT profiles_pkey PRIMARY KEY (id),
  CONSTRAINT profiles_id_fkey FOREIGN KEY (id) REFERENCES auth.users(id)
);
CREATE TABLE public.doctors (
  id uuid NOT NULL DEFAULT gen_random_uuid(),
  name text NOT NULL UNIQUE,
  specialty text,
  active boolean NOT NULL DEFAULT true,
  created_at timestamp with time zone NOT NULL DEFAULT now(),
  CONSTRAINT doctors_pkey PRIMARY KEY (id)
);
CREATE TABLE public.exam_types (
  id uuid NOT NULL DEFAULT gen_random_uuid(),
  name text NOT NULL UNIQUE,
  active boolean NOT NULL DEFAULT true,
  created_at timestamp with time zone NOT NULL DEFAULT now(),
  CONSTRAINT exam_types_pkey PRIMARY KEY (id)
);
CREATE TABLE public.procedures (
  id uuid NOT NULL DEFAULT gen_random_uuid(),
  name text NOT NULL UNIQUE,
  category text,
  active boolean NOT NULL DEFAULT true,
  created_at timestamp with time zone NOT NULL DEFAULT now(),
  CONSTRAINT procedures_pkey PRIMARY KEY (id)
);
CREATE TABLE public.insurances (
  id uuid NOT NULL DEFAULT gen_random_uuid(),
  name text NOT NULL UNIQUE,
  active boolean NOT NULL DEFAULT true,
  created_at timestamp with time zone NOT NULL DEFAULT now(),
  CONSTRAINT insurances_pkey PRIMARY KEY (id)
);
CREATE TABLE public.patients (
  id uuid NOT NULL DEFAULT gen_random_uuid(),
  full_name text NOT NULL,
  whatsapp text,
  birth_date date,
  notes text,
  created_by uuid,
  created_at timestamp with time zone NOT NULL DEFAULT now(),
  updated_at timestamp with time zone NOT NULL DEFAULT now(),
  CONSTRAINT patients_pkey PRIMARY KEY (id),
  CONSTRAINT patients_created_by_fkey FOREIGN KEY (created_by) REFERENCES public.profiles(id)
);
CREATE TABLE public.exams (
  id uuid NOT NULL DEFAULT gen_random_uuid(),
  exam_date date NOT NULL,
  patient_id uuid NOT NULL,
  doctor_id uuid NOT NULL,
  exam_type_id uuid NOT NULL,
  status USER-DEFINED NOT NULL DEFAULT 'Agendado'::exam_status,
  observation text,
  launched_by uuid NOT NULL,
  created_at timestamp with time zone NOT NULL DEFAULT now(),
  updated_at timestamp with time zone NOT NULL DEFAULT now(),
  launched_by_name text,
  launch_group_id uuid NOT NULL,
  CONSTRAINT exams_pkey PRIMARY KEY (id),
  CONSTRAINT exams_patient_id_fkey FOREIGN KEY (patient_id) REFERENCES public.patients(id),
  CONSTRAINT exams_doctor_id_fkey FOREIGN KEY (doctor_id) REFERENCES public.doctors(id),
  CONSTRAINT exams_exam_type_id_fkey FOREIGN KEY (exam_type_id) REFERENCES public.exam_types(id),
  CONSTRAINT exams_launched_by_fkey FOREIGN KEY (launched_by) REFERENCES public.profiles(id)
);
CREATE TABLE public.surgeries (
  id uuid NOT NULL DEFAULT gen_random_uuid(),
  surgery_date date NOT NULL,
  patient_id uuid NOT NULL,
  doctor_id uuid NOT NULL,
  procedure_id uuid NOT NULL,
  eye text NOT NULL DEFAULT 'Não se aplica'::text CHECK (eye = ANY (ARRAY['Não se aplica'::text, 'Direito (OD)'::text, 'Esquerdo (OE)'::text, 'Ambos'::text])),
  insurance_id uuid,
  status USER-DEFINED NOT NULL DEFAULT 'Solicitação'::surgery_status,
  arrival_time time without time zone,
  surgery_time time without time zone,
  payment_status text DEFAULT 'Não informado'::text,
  observation text,
  launched_by uuid NOT NULL,
  created_at timestamp with time zone NOT NULL DEFAULT now(),
  updated_at timestamp with time zone NOT NULL DEFAULT now(),
  launched_by_name text,
  launch_group_id uuid NOT NULL,
  CONSTRAINT surgeries_pkey PRIMARY KEY (id),
  CONSTRAINT surgeries_patient_id_fkey FOREIGN KEY (patient_id) REFERENCES public.patients(id),
  CONSTRAINT surgeries_doctor_id_fkey FOREIGN KEY (doctor_id) REFERENCES public.doctors(id),
  CONSTRAINT surgeries_procedure_id_fkey FOREIGN KEY (procedure_id) REFERENCES public.procedures(id),
  CONSTRAINT surgeries_insurance_id_fkey FOREIGN KEY (insurance_id) REFERENCES public.insurances(id),
  CONSTRAINT surgeries_launched_by_fkey FOREIGN KEY (launched_by) REFERENCES public.profiles(id)
);