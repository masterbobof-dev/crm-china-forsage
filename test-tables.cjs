const { createClient } = require('@supabase/supabase-js');
const supabaseUrl = 'https://yecsylvpezwpuyomureg.supabase.co';
const supabaseAnonKey = 'sb_publishable_QyrZOkrM1rusdcovMNLL7w_pk5Ta0xM';
const supabase = createClient(supabaseUrl, supabaseAnonKey);

async function test() {
  const tables = ['purchases', 'batches', 'local_deliveries', 'settings'];
  for (const table of tables) {
    const { error } = await supabase.from(table).select('*').limit(1);
    if (error) {
      console.log(`Table ${table} error:`, error.message);
    } else {
      console.log(`Table ${table} exists.`);
    }
  }
}
test();
