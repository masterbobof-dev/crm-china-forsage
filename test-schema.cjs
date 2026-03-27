const { createClient } = require('@supabase/supabase-js');
const supabaseUrl = 'https://yecsylvpezwpuyomureg.supabase.co';
const supabaseAnonKey = 'sb_publishable_QyrZOkrM1rusdcovMNLL7w_pk5Ta0xM';
const supabase = createClient(supabaseUrl, supabaseAnonKey);

async function test() {
  const { data: pData } = await supabase.from('purchases').select('*').limit(1);
  console.log('purchases schema:', pData);
  const { data: bData } = await supabase.from('batches').select('*').limit(1);
  console.log('batches schema:', bData);
}
test();
