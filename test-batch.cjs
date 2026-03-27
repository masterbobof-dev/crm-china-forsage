const { createClient } = require('@supabase/supabase-js');
const supabaseUrl = 'https://yecsylvpezwpuyomureg.supabase.co';
const supabaseAnonKey = 'sb_publishable_QyrZOkrM1rusdcovMNLL7w_pk5Ta0xM';
const supabase = createClient(supabaseUrl, supabaseAnonKey);

async function test() {
  const { error } = await supabase.from('batches').insert([{ id: 'test_batch', name: 'test' }]);
  console.log('Insert error:', error);
  const { data } = await supabase.from('batches').select('*').eq('id', 'test_batch');
  console.log('Row:', data);
}
test();
