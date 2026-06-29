const tabs=document.querySelectorAll('.tab');const panels=document.querySelectorAll('.panel');tabs.forEach(tab=>tab.addEventListener('click',()=>{tabs.forEach(t=>t.classList.remove('active'));panels.forEach(p=>p.classList.remove('active'));tab.classList.add('active');document.getElementById(tab.dataset.tab).classList.add('active');window.scrollTo({top:0,behavior:'smooth'});}));document.querySelectorAll('.payment').forEach(btn=>btn.addEventListener('click',()=>{const statuses=['Pending','Partial','Paid'];let next=statuses[(statuses.indexOf(btn.dataset.status)+1)%statuses.length];btn.dataset.status=next;btn.textContent=next;btn.className='payment '+next.toLowerCase();}));const toast=document.getElementById('toast');document.querySelectorAll('.copy-btn').forEach(btn=>btn.addEventListener('click',async()=>{try{await navigator.clipboard.writeText(btn.dataset.copy);}catch(e){}toast.classList.add('show');setTimeout(()=>toast.classList.remove('show'),1000);}));

/* v7.3 navigation fixes */
function activateTab(tabName) {
  const targetTab = document.querySelector(`.tab[data-tab="${tabName}"]`);
  const targetPanel = document.getElementById(tabName);
  if (!targetTab || !targetPanel) return;

  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));

  targetTab.classList.add('active');
  targetPanel.classList.add('active');
}

window.addEventListener('load', () => {
  activateTab('home');
  setTimeout(() => window.scrollTo({ top: 0, left: 0, behavior: 'auto' }), 0);
});

document.querySelectorAll('[data-target-tab][data-scroll-target]').forEach(el => {
  el.addEventListener('click', () => {
    const tabName = el.dataset.targetTab;
    const scrollTarget = el.dataset.scrollTarget;

    activateTab(tabName);

    setTimeout(() => {
      const target = document.getElementById(scrollTarget);
      if (target) {
        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
        target.classList.add('highlight');
        setTimeout(() => target.classList.remove('highlight'), 1600);
      }
    }, 80);
  });
});
