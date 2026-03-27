const { createClient } = require('@supabase/supabase-js');
const supabaseUrl = 'https://yecsylvpezwpuyomureg.supabase.co';
const supabaseAnonKey = 'sb_publishable_QyrZOkrM1rusdcovMNLL7w_pk5Ta0xM';
const supabase = createClient(supabaseUrl, supabaseAnonKey);

async function test() {
  const { error } = await supabase.from('purchases').insert([{ 
    id: 'test', 
    name: 'test',
    price_yuan: 100,
    quantity: 1,
    exchange_rate: 5.5,
    track_number: '123',
    status: 'purchased'
  }]);
  console.log('Insert error:', error);
}
test();
