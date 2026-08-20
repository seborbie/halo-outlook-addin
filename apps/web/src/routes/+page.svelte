<svelte:head>
  <title>InboxLink — Outlook meets HaloPSA</title>
  <meta name="description" content="Attach Outlook conversations to HaloPSA tickets, keep replies in sync, and protect every customer with strict tenant isolation." />
</svelte:head>

<section class="hero">
  <div class="shell hero-grid">
    <div class="hero-copy">
      <p class="eyebrow">The Outlook add-in for HaloPSA teams</p>
      <h1>Close the gap between <span>inbox and ticket.</span></h1>
      <p class="lede">Attach the right email chain, find the right Halo ticket, and keep every reply in sync—without leaving Outlook.</p>
      <div class="hero-actions">
        <a class="button" href="/signup">Start a 14-day pilot <span aria-hidden="true">→</span></a>
        <a class="button button-secondary" href="/docs/overview">Read the setup guide</a>
      </div>
      <div class="proof-row">
        <span><i class="proof-dot"></i> No card required</span>
        <span>Microsoft Entra sign-in</span>
        <span>Built for multi-company rollout</span>
      </div>
    </div>

    <div class="product-window" aria-label="InboxLink add-in preview">
      <div class="window-bar"><i></i><i></i><i></i></div>
      <div class="window-body">
        <div class="mock-header">
          <span class="mock-title">InboxLink</span>
          <span class="connected">● Connected</span>
        </div>
        <div class="email-card">
          <small>RE: New starter setup</small>
          <h4>From Mia Chen · Northstar Design</h4>
          <div class="email-lines"><i></i><i></i><i></i></div>
        </div>
        <div class="ticket-panel">
          <div class="search-box"><span aria-hidden="true">⌕</span> Search ticket number or subject</div>
          <div class="ticket">
            <span class="ticket-id">4821</span>
            <span><strong>New starter setup</strong><small>Northstar Design · In progress</small></span>
            <button type="button">Attach</button>
          </div>
          <div class="ticket">
            <span class="ticket-id">4794</span>
            <span><strong>Microsoft 365 onboarding</strong><small>Northstar Design · Open</small></span>
            <button type="button">Attach</button>
          </div>
        </div>
      </div>
    </div>
  </div>
</section>

<div class="logo-strip">
  <div class="shell logo-strip-inner">
    <span>Works where your team already works</span>
    <strong>Microsoft Outlook</strong>
    <span>＋</span>
    <strong>HaloPSA</strong>
  </div>
</div>

<section class="section" id="features">
  <div class="shell">
    <div class="section-heading">
      <p class="eyebrow">Less tab switching, better tickets</p>
      <h2>The email workflow your service desk has been missing.</h2>
      <p class="lede">InboxLink keeps context intact from the first customer message to the final reply.</p>
    </div>
    <div class="feature-grid">
      <article class="feature-card accent">
        <span class="feature-number">01</span>
        <h3>Attach the full chain once</h3>
        <p>Choose a Halo ticket and capture the conversation history on the first attach—formatted and readable.</p>
      </article>
      <article class="feature-card">
        <span class="feature-number">02</span>
        <h3>Sync only new replies</h3>
        <p>InboxLink remembers the thread, avoids duplicates, and adds later replies to the same ticket automatically.</p>
      </article>
      <article class="feature-card">
        <span class="feature-number">03</span>
        <h3>Find tickets in seconds</h3>
        <p>See assigned open tickets or search by ticket number without leaving the message you are working on.</p>
      </article>
      <article class="feature-card">
        <span class="feature-number">04</span>
        <h3>Secure OAuth connections</h3>
        <p>Microsoft identifies the user; encrypted Halo grants keep passwords and client secrets out of the add-in.</p>
      </article>
      <article class="feature-card">
        <span class="feature-number">05</span>
        <h3>Classic and new Outlook</h3>
        <p>Use the task pane for received mail and background send events to keep outgoing replies attached.</p>
      </article>
      <article class="feature-card">
        <span class="feature-number">06</span>
        <h3>Company-by-company isolation</h3>
        <p>Every session, ticket mapping, and Halo connection is scoped to one Microsoft organisation at the database layer.</p>
      </article>
    </div>
  </div>
</section>

<section class="workflow-section">
  <div class="shell workflow-grid">
    <div>
      <p class="eyebrow">Simple by design</p>
      <h2>Three clicks from email to Halo.</h2>
      <p class="lede">No copy-paste. No hunting through browser tabs. No mystery about whether the ticket has the latest reply.</p>
    </div>
    <div class="steps">
      <article class="step"><span class="step-number">1</span><div><h3>Open an email</h3><p>InboxLink reads only the message data needed for the action you choose.</p></div></article>
      <article class="step"><span class="step-number">2</span><div><h3>Choose the Halo ticket</h3><p>Pick from your open queue or search for an exact ticket number.</p></div></article>
      <article class="step"><span class="step-number">3</span><div><h3>Keep the conversation moving</h3><p>Later replies are matched to the original ticket and duplicates are skipped.</p></div></article>
    </div>
  </div>
</section>

<section class="section">
  <div class="shell security-callout">
    <div>
      <p class="eyebrow">SaaS-ready foundations</p>
      <h2>One product. Hard walls between companies.</h2>
      <p>Microsoft tenant identity is carried through users, encrypted Halo grants, browser sessions, bug-report links, and message mappings. PostgreSQL row-level security enforces the same boundary even if an application query is wrong.</p>
    </div>
    <div class="security-badge" aria-hidden="true">⌁</div>
  </div>
</section>

<section class="cta">
  <div class="shell cta-card">
    <div><h2>Give your service desk its focus back.</h2><p>Start with one Microsoft organisation and scale when your team is ready.</p></div>
    <a class="button" href="/signup">Start your pilot <span aria-hidden="true">→</span></a>
  </div>
</section>
