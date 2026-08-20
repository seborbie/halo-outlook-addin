export type DocEntry = {
  slug: string;
  title: string;
  description: string;
  section: string;
  order: number;
};

export const docs: DocEntry[] = [
  {
    slug: "overview",
    title: "Overview",
    description: "How InboxLink fits between Outlook, Microsoft Entra, and HaloPSA.",
    section: "Get started",
    order: 1,
  },
  {
    slug: "quickstart",
    title: "Quickstart",
    description: "Get a pilot tenant running from a clean deployment.",
    section: "Get started",
    order: 2,
  },
  {
    slug: "microsoft-entra",
    title: "Microsoft Entra setup",
    description: "Register the multi-tenant app and grant the required delegated scope.",
    section: "Configure",
    order: 3,
  },
  {
    slug: "halopsa",
    title: "HaloPSA API setup",
    description: "Create the native OAuth application used by each customer.",
    section: "Configure",
    order: 4,
  },
  {
    slug: "postgres",
    title: "PostgreSQL deployment",
    description: "Provision the database, apply migrations, and configure secure connections.",
    section: "Deploy",
    order: 5,
  },
  {
    slug: "tenant-onboarding",
    title: "Tenant onboarding",
    description: "Create an organisation and connect its own HaloPSA instance.",
    section: "Operate",
    order: 6,
  },
  {
    slug: "security",
    title: "Security and isolation",
    description: "Understand tenant boundaries, encrypted grants, and retained data.",
    section: "Operate",
    order: 7,
  },
  {
    slug: "troubleshooting",
    title: "Troubleshooting",
    description: "Resolve sign-in, callback, database, and add-in loading problems.",
    section: "Help",
    order: 8,
  },
];

export const docsBySection = docs.reduce<Record<string, DocEntry[]>>((groups, doc) => {
  groups[doc.section] ||= [];
  groups[doc.section].push(doc);
  return groups;
}, {});
