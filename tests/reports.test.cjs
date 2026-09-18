const {test}=require('node:test');
const assert=require('node:assert/strict');
const R=require('../reports.js');
const rows=[
 {id:'b',receipt_date:'2026-09-01',vendor:'Wash',amount:'10.10',currency:'USD',company:'Wash',status:'ready',reimbursable:true},
 {id:'a',receipt_date:'2026-09-01',vendor:'Wash',amount:'-10.10',currency:'USD',company:'Wash',status:'ready',reimbursable:true},
 {id:'c',receipt_date:'2026-09-02',vendor:'École',amount:'3.22',currency:'CAD',company:'SMI',status:'ready'},
 {id:'d',receipt_date:'2026-09-03',vendor:'Excluded',amount:'99.99',currency:'USD',status:'excluded'},
 {id:'e',receipt_date:'2026-08-01',vendor:'Old',amount:'1.00',currency:'USD',status:'ready'}
];
test('report snapshot applies filters, stable order, signed cents and separate currencies',()=>{
 const r=R.snapshot(rows,{month:'2026-09'});
 assert.deepEqual(r.rows.map(x=>x.id),['a','b','c']);
 assert.equal(r.omitted,1);
 assert.deepEqual(r.totals,[{currency:'CAD',count:1,total:3.22,reimbursable:0,excludedCount:0,excluded:0},{currency:'USD',count:2,total:0,reimbursable:0,excludedCount:0,excluded:0}]);
 assert.equal(R.snapshot(rows,{month:'2026-09',status:'excluded'}).rows.length,0);
 const inc=R.snapshot(rows,{month:'2026-09',includeExcluded:true});
 assert.equal(inc.totals[1].total,0); assert.equal(inc.totals[1].excluded,99.99);
 assert.equal(R.snapshot(rows,{company:'Wash',search:'wash',status:'ready'}).rows.length,2);
 assert.equal(R.snapshot(rows,{month:'2000-01'}).rows.length,0);
 assert.throws(()=>R.snapshot([{...rows[0],amount:null}],{}),/amount/);
});
