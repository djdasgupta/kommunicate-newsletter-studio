export default async (request) => {
  if (request.method !== 'POST') return new Response(JSON.stringify({error:'Method not allowed'}),{status:405});
  try {
    const {to,subject,message,html,reviewLink}=await request.json();
    if(!Array.isArray(to)||!to.length||!html) return new Response(JSON.stringify({error:'Recipients and newsletter HTML are required'}),{status:400});
    const apiKey=process.env.RESEND_API_KEY;
    const from=process.env.NEWSLETTER_FROM;
    if(!apiKey||!from) return new Response(JSON.stringify({error:'Email service is not configured'}),{status:500});
    const intro=`<div style="font-family:Arial,sans-serif;max-width:740px;margin:0 auto 20px;padding:18px;border:1px solid #e5e7eb;border-radius:12px"><p style="margin:0 0 10px"><strong>Review request</strong></p><p style="margin:0 0 10px">${escapeHtml(message||'Please review this newsletter.')}</p>${reviewLink?`<p style="margin:0"><a href="${reviewLink}">Open interactive review version</a></p>`:''}</div>`;
    const result=await fetch('https://api.resend.com/emails',{method:'POST',headers:{Authorization:`Bearer ${apiKey}`,'Content-Type':'application/json'},body:JSON.stringify({from,to,subject:subject||'Newsletter review',html:intro+html})});
    const data=await result.json();
    if(!result.ok) return new Response(JSON.stringify({error:data.message||'Email provider rejected the request'}),{status:result.status});
    return new Response(JSON.stringify({ok:true,id:data.id}),{status:200,headers:{'Content-Type':'application/json'}});
  } catch (error) {return new Response(JSON.stringify({error:error.message}),{status:500,headers:{'Content-Type':'application/json'}})}
};
function escapeHtml(value=''){return String(value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
