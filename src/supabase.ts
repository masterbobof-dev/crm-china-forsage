import { createClient } from '@supabase/supabase-js';

// @ts-ignore
const supabaseUrl = import.meta.env.VITE_SUPABASE_URL || 'https://yecsylvpezwpuyomureg.supabase.co';
// @ts-ignore
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY || 'sb_publishable_QyrZOkrM1rusdcovMNLL7w_pk5Ta0xM';

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
