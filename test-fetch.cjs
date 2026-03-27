const { createClient } = require('@supabase/supabase-js');
const supabaseUrl = 'https://yecsylvpezwpuyomureg.supabase.co';
const supabaseAnonKey = 'sb_publishable_QyrZOkrM1rusdcovMNLL7w_pk5Ta0xM';
const supabase = createClient(supabaseUrl, supabaseAnonKey);

async function test() {
  const { data, error } = await supabase.from('purchases').select('*').eq('id', 'test');
  console.log('Row:', data);
}
test();
