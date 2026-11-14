import { DOC_VERSION, SLIDE_W, type Presentation, type Slide, type SlideElement, type Transition } from './types';
import { THEMES, getTheme, type Theme } from './themes';
import { makeShape, makeSlide, makeText, type LayoutId } from './factory';
import { parasFromText } from './richtext';
import { t } from './i18n';

/** A ready-made starting point shown on the start screen. */
export interface Template {
  id: string;
  name: string;
  description: string;
  kind: 'template' | 'theme';
  /** Language of the content, shown as a tag on the card ('ar' Arabic, 'ar-en' bilingual). */
  lang?: 'ar' | 'ar-en';
  build: () => Presentation;
}

interface SlideSpec {
  layout: LayoutId;
  title?: string;
  /** One string per body placeholder; lines become bullets on content layouts. */
  body?: string[];
  notes?: string;
  /** Big-number cards laid out under the title. */
  stats?: [value: string, label: string][];
}

function bar(theme: Theme, x: number, y: number, w: number, h: number): SlideElement {
  return { ...makeShape(theme, 'rect'), x, y, w, h, stroke: null, fill: theme.accent };
}

function statCards(theme: Theme, stats: [string, string][], rtl: boolean): SlideElement[] {
  const w = 260;
  const gap = (SLIDE_W - 120 - stats.length * w) / Math.max(1, stats.length - 1);
  // Right-to-left decks read the cards from the right.
  return (rtl ? [...stats].reverse() : stats).flatMap(([value, label], i) => {
    const x = 60 + i * (w + gap);
    const card = { ...makeShape(theme, 'roundRect'), x, y: 170, w, h: 160, stroke: null, fill: theme.accent };
    card.paras = parasFromText(value);
    card.style = { ...card.style, fontFamily: theme.titleFont, fontSize: 54, bold: true };
    const caption = makeText(theme, { x, y: 345, w, h: 50, text: label, role: 'body', style: { fontSize: 20, align: 'center' } });
    return [card, caption];
  });
}

const mirror = (el: SlideElement | null) => el && { ...el, x: SLIDE_W - el.x - el.w };

function buildSlide(theme: Theme, spec: SlideSpec, transition: Transition, decorate: boolean, rtl = false): Slide {
  const slide = makeSlide(theme, spec.stats ? 'titleOnly' : spec.layout);
  // Right-to-left decks fill the right-hand column first.
  const body = rtl && spec.layout === 'twoContent' ? [...(spec.body ?? [])].reverse() : (spec.body ?? []);
  const texts = [spec.title, ...body];
  slide.elements.forEach((el, i) => {
    if (texts[i] !== undefined) el.paras = parasFromText(texts[i]!);
  });
  if (decorate) {
    const layout = spec.stats ? 'titleOnly' : spec.layout;
    const deco =
      layout === 'title'
        ? bar(theme, (SLIDE_W - 100) / 2, 286, 100, 5)
        : layout === 'section'
          ? bar(theme, 80, 302, 100, 5)
          : layout === 'blank'
            ? null
            : bar(theme, 60, 118, 70, 5);
    if (deco) slide.elements.unshift(rtl ? mirror(deco)! : deco);
  }
  if (spec.stats) slide.elements.push(...statCards(theme, spec.stats, rtl));
  slide.notes = spec.notes ?? '';
  slide.transition = transition;
  return slide;
}

function deck(title: string, themeId: string, specs: SlideSpec[], transition: Transition = 'fade', rtl = false): Presentation {
  const theme = getTheme(themeId);
  return { version: DOC_VERSION, title, themeId: theme.id, slides: specs.map((s) => buildSlide(theme, s, transition, true, rtl)) };
}

const lines = (...l: string[]) => l.join('\n');

/** Arabic and bilingual (Arabic/English) templates. Bilingual slides put Arabic on the right, English on the left. */
const ARABIC_TEMPLATES: Template[] = [
  {
    id: 'ar-company',
    name: 'Company Profile (Arabic/English)',
    description: 'A bilingual company profile: about us, vision, services, key figures and contact, in Arabic and English side by side.',
    kind: 'template',
    lang: 'ar-en',
    build: () =>
      deck(
        'الملف التعريفي · Company Profile',
        'diwan',
        [
          { layout: 'title', title: lines('اسم الشركة', 'Company Name'), body: [lines('شعار الشركة في سطر واحد', 'Your tagline in one line')] },
          {
            layout: 'twoContent',
            title: 'من نحن · About Us',
            body: [
              lines('تأسست الشركة عام ٢٠١٠ في الرياض', 'نقدّم حلولاً متكاملة للقطاعين العام والخاص', 'فريق من ١٢٠ موظفاً في ٣ مدن'),
              lines('Founded in 2010 in Riyadh', 'Complete solutions for the public and private sectors', 'A team of 120 people in 3 cities'),
            ],
            notes: 'اجعل العمودين يقولان الشيء نفسه، ليصل المحتوى كاملاً إلى كل قارئ.',
          },
          {
            layout: 'twoContent',
            title: 'رؤيتنا ورسالتنا · Vision & Mission',
            body: [
              lines('الرؤية: أن نكون الشريك الأول لعملائنا في المنطقة', 'الرسالة: خدمات موثوقة بجودة عالمية'),
              lines('Vision: to be our clients’ first-choice partner in the region', 'Mission: dependable services to world-class standards'),
            ],
          },
          {
            layout: 'twoContent',
            title: 'خدماتنا · Our Services',
            body: [lines('الاستشارات', 'التصميم والتنفيذ', 'التدريب', 'الدعم والصيانة'), lines('Consulting', 'Design and delivery', 'Training', 'Support and maintenance')],
          },
          {
            layout: 'titleOnly',
            title: 'أرقامنا · Key Figures',
            stats: [
              ['+٥٠٠', 'مشروع منجز · Projects'],
              ['١٤', 'عاماً من الخبرة · Years'],
              ['٩٨٪', 'رضا العملاء · Satisfaction'],
            ],
          },
          {
            layout: 'twoContent',
            title: 'عملاؤنا · Our Clients',
            body: [lines('الجهات الحكومية', 'البنوك والمؤسسات المالية', 'قطاع التجزئة'), lines('Government entities', 'Banks and financial institutions', 'Retail')],
          },
          {
            layout: 'section',
            title: 'تواصل معنا · Contact Us',
            body: [lines('info@company.com · ‎+966 11 000 0000', 'www.company.com')],
          },
        ],
        'fade',
        true,
      ),
  },
  {
    id: 'ar-pitch',
    name: 'Pitch Deck (Arabic)',
    description: 'An Arabic investor pitch: problem, solution, market, traction, business model, team and the ask.',
    kind: 'template',
    lang: 'ar',
    build: () =>
      deck(
        'العرض على المستثمرين',
        'sahara',
        [
          { layout: 'title', title: 'اسم الشركة', body: ['ما نقدّمه في جملة واحدة'] },
          {
            layout: 'titleContent',
            title: 'المشكلة',
            body: [lines('من يعاني من المشكلة، وما هي؟', 'كيف يتعاملون معها اليوم', 'لماذا لا يكفي ذلك')],
            notes: 'اجعل المشكلة ملموسة: قصة قصيرة عن عميل حقيقي تنجح هنا.',
          },
          {
            layout: 'titleContent',
            title: 'الحل',
            body: [lines('ما الذي بنيناه، في جملة واحدة', 'كيف يحلّ المشكلة', 'ما الذي يصعّب تقليده')],
          },
          {
            layout: 'twoContent',
            title: 'حجم السوق',
            body: [lines('السوق الكلي: X مليار', 'السوق القابل للخدمة: Y مليار', 'السوق المستهدف: Z مليون'), lines('الاتجاهات الداعمة للنمو', 'لماذا الآن؟')],
          },
          {
            layout: 'titleOnly',
            title: 'الإنجازات حتى الآن',
            stats: [
              ['١٠ آلاف', 'مستخدم نشط'],
              ['٢٥٠ ألف $', 'إيرادات سنوية متكررة'],
              ['٣٥٪', 'نمو شهري'],
            ],
          },
          {
            layout: 'titleContent',
            title: 'نموذج العمل',
            body: [lines('مصادر الإيرادات', 'التسعير واقتصاديات الوحدة', 'الطريق إلى الربحية')],
          },
          {
            layout: 'twoContent',
            title: 'الفريق',
            body: [lines('اسم المؤسس، الرئيس التنفيذي', 'الخبرة ذات الصلة'), lines('اسم المؤسس، المدير التقني', 'الخبرة ذات الصلة')],
          },
          { layout: 'section', title: 'ما نطلبه', body: ['نسعى لجمع X لتحقيق [الهدف] بحلول [التاريخ]'] },
        ],
        'fade',
        true,
      ),
  },
  {
    id: 'ar-lesson',
    name: 'Classroom Lesson (Arabic)',
    description: 'An Arabic lesson plan: objectives, new vocabulary, explanation, an activity, review questions and homework.',
    kind: 'template',
    lang: 'ar',
    build: () =>
      deck(
        'درس',
        'andalus',
        [
          { layout: 'title', title: 'عنوان الدرس', body: ['المادة · الصف · التاريخ'] },
          {
            layout: 'titleContent',
            title: 'أهداف الدرس',
            body: [lines('في نهاية الدرس سيكون الطالب قادراً على:', 'شرح الفكرة الرئيسية', 'تطبيقها على مثال جديد', 'ربطها بما تعلّمه سابقاً')],
          },
          {
            layout: 'twoContent',
            title: 'المفردات الجديدة',
            body: [lines('الكلمة الأولى: معناها', 'الكلمة الثانية: معناها', 'الكلمة الثالثة: معناها'), lines('مثال في جملة', 'مثال في جملة', 'مثال في جملة')],
          },
          {
            layout: 'titleContent',
            title: 'الشرح',
            body: [lines('الفكرة الأولى مع مثال', 'الفكرة الثانية مع مثال', 'خلاصة في جملة واحدة')],
            notes: 'توقّف بعد كل فكرة واسأل سؤالاً قصيراً للتأكد من الفهم.',
          },
          {
            layout: 'section',
            title: 'نشاط',
            body: ['اعملوا في مجموعات من ٣ طلاب لمدة ١٠ دقائق'],
          },
          {
            layout: 'titleContent',
            title: 'أسئلة المراجعة',
            body: [lines('ما المقصود بـ ...؟', 'اذكر مثالاً على ...', 'لماذا ...؟')],
          },
          {
            layout: 'titleContent',
            title: 'الواجب المنزلي',
            body: [lines('حلّ التمارين ١ إلى ٥ صفحة ٤٢', 'التسليم: يوم الأحد')],
          },
        ],
        'fade',
        true,
      ),
  },
  {
    id: 'ar-event',
    name: 'Event / Conference (Arabic/English)',
    description: 'Bilingual event slides: welcome, agenda, speakers, sponsors and venue information, in Arabic and English.',
    kind: 'template',
    lang: 'ar-en',
    build: () =>
      deck(
        'المؤتمر · Conference',
        'nile',
        [
          {
            layout: 'title',
            title: lines('اسم المؤتمر ٢٠٢٦', 'Conference Name 2026'),
            body: [lines('١٥–١٦ نوفمبر · المدينة', '15–16 November · City')],
          },
          {
            layout: 'twoContent',
            title: 'جدول الأعمال · Agenda',
            body: [
              lines('٩:٠٠  التسجيل والقهوة', '١٠:٠٠  الكلمة الافتتاحية', '١١:٠٠  الجلسة الأولى', '١:٠٠  الغداء', '٢:٠٠  ورش العمل'),
              lines('9:00  Registration and coffee', '10:00  Opening keynote', '11:00  First session', '1:00  Lunch', '2:00  Workshops'),
            ],
          },
          {
            layout: 'twoContent',
            title: 'المتحدثون · Speakers',
            body: [lines('د. اسم المتحدث – المنصب', 'أ. اسم المتحدث – المنصب'), lines('Dr. Speaker Name – Title', 'Speaker Name – Title')],
          },
          {
            layout: 'twoContent',
            title: 'الرعاة · Sponsors',
            body: [lines('الراعي البلاتيني', 'الراعي الذهبي', 'الراعي الفضي'), lines('Platinum sponsor', 'Gold sponsor', 'Silver sponsor')],
          },
          {
            layout: 'twoContent',
            title: 'معلومات عملية · Venue & Info',
            body: [lines('المكان: قاعة المؤتمرات', 'شبكة Wi-Fi: Event2026', 'التطبيق: حمّل تطبيق المؤتمر'), lines('Venue: Conference Hall', 'Wi-Fi: Event2026', 'App: download the conference app')],
          },
          { layout: 'section', title: 'شكراً لحضوركم · Thank You', body: ['#Conference2026'] },
        ],
        'push',
        true,
      ),
  },
];

export const TEMPLATES: Template[] = [
  {
    id: 'blank',
    name: 'Blank Presentation',
    description: 'Start from scratch with a single title slide.',
    kind: 'template',
    build: () => {
      const theme = getTheme('office');
      return { version: DOC_VERSION, title: t('Untitled presentation'), themeId: theme.id, slides: [makeSlide(theme, 'title')] };
    },
  },
  {
    id: 'pitch',
    name: 'Pitch Deck',
    description: 'Tell investors your story: problem, solution, market, traction, team and the ask.',
    kind: 'template',
    build: () =>
      deck('Pitch Deck', 'midnight', [
        { layout: 'title', title: 'Company Name', body: ['The one-line pitch for what you do'] },
        {
          layout: 'titleContent',
          title: 'The Problem',
          body: [lines('Who is struggling, and with what?', 'How they solve it today', 'Why that isn’t good enough')],
          notes: 'Make the pain concrete. A short customer story works well here.',
        },
        {
          layout: 'titleContent',
          title: 'Our Solution',
          body: [lines('What you built, in one sentence', 'How it solves the problem', 'What makes it hard to copy')],
        },
        {
          layout: 'twoContent',
          title: 'Market Opportunity',
          body: [
            lines('Total market: $X B', 'Serviceable market: $Y B', 'Target market: $Z M'),
            lines('Trends driving growth', 'Why now'),
          ],
        },
        {
          layout: 'titleOnly',
          title: 'Traction',
          stats: [
            ['10k', 'Active users'],
            ['$250k', 'Annual recurring revenue'],
            ['35%', 'Month-over-month growth'],
          ],
        },
        {
          layout: 'titleContent',
          title: 'Business Model',
          body: [lines('How you make money', 'Pricing and unit economics', 'Path to profitability')],
        },
        {
          layout: 'twoContent',
          title: 'The Team',
          body: [lines('Founder name, CEO', 'Relevant experience'), lines('Founder name, CTO', 'Relevant experience')],
        },
        { layout: 'section', title: 'The Ask', body: ['Raising $X to reach [milestone] by [date]'] },
      ]),
  },
  {
    id: 'status',
    name: 'Project Status Report',
    description: 'Share progress, risks and next steps with stakeholders.',
    kind: 'template',
    build: () =>
      deck('Project Status Report', 'office', [
        { layout: 'title', title: 'Project Status Report', body: ['Project name · Reporting period'] },
        {
          layout: 'titleContent',
          title: 'Summary',
          body: [lines('Overall status: On track', 'Key accomplishments this period', 'Upcoming milestones')],
        },
        {
          layout: 'titleOnly',
          title: 'At a Glance',
          stats: [
            ['85%', 'Complete'],
            ['12', 'Tasks finished'],
            ['2', 'Open risks'],
          ],
        },
        {
          layout: 'twoContent',
          title: 'Progress',
          body: [lines('Completed', 'Milestone one', 'Milestone two'), lines('In progress', 'Milestone three', 'Milestone four')],
        },
        {
          layout: 'titleContent',
          title: 'Risks and Issues',
          body: [lines('Risk: description · Owner · Mitigation', 'Issue: description · Owner · Due date')],
        },
        {
          layout: 'titleContent',
          title: 'Next Steps',
          body: [lines('Action item · Owner · Date', 'Action item · Owner · Date', 'Decisions needed')],
        },
      ]),
  },
  {
    id: 'qbr',
    name: 'Quarterly Business Review',
    description: 'Review the quarter’s results and set priorities for the next one.',
    kind: 'template',
    build: () =>
      deck('Quarterly Business Review', 'ocean', [
        { layout: 'title', title: 'Q3 Business Review', body: ['Team name · Date'] },
        {
          layout: 'titleOnly',
          title: 'Highlights',
          stats: [
            ['$4.2M', 'Revenue'],
            ['+18%', 'Growth vs. last quarter'],
            ['96%', 'Customer retention'],
          ],
        },
        {
          layout: 'titleContent',
          title: 'Key Wins',
          body: [lines('Biggest win of the quarter', 'New customers and partnerships', 'Launches and improvements')],
        },
        {
          layout: 'twoContent',
          title: 'Lessons Learned',
          body: [lines('What worked', 'Keep doing this'), lines('What didn’t', 'Change this next quarter')],
        },
        { layout: 'section', title: 'Looking Ahead', body: ['Plans for Q4'] },
        {
          layout: 'titleContent',
          title: 'Q4 Priorities',
          body: [lines('Priority one and how we’ll measure it', 'Priority two', 'Priority three')],
        },
      ]),
  },
  {
    id: 'launch',
    name: 'Product Launch',
    description: 'Introduce a new product: what it is, its features and when it ships.',
    kind: 'template',
    build: () =>
      deck(
        'Product Launch',
        'sunset',
        [
          { layout: 'title', title: 'Introducing Product', body: ['A tagline that sticks'] },
          {
            layout: 'titleContent',
            title: 'What It Is',
            body: [lines('The product in one sentence', 'Who it’s for', 'The problem it solves')],
          },
          {
            layout: 'twoContent',
            title: 'Key Features',
            body: [lines('Feature one', 'Feature two', 'Feature three'), lines('Feature four', 'Feature five', 'Feature six')],
          },
          {
            layout: 'titleContent',
            title: 'Pricing and Availability',
            body: [lines('Plans and prices', 'Launch date', 'Where to get it')],
          },
          { layout: 'section', title: 'Thank You', body: ['Questions?'] },
        ],
        'zoom',
      ),
  },
  {
    id: 'lesson',
    name: 'Classroom Lesson',
    description: 'Plan a lesson with objectives, vocabulary, an activity and homework.',
    kind: 'template',
    build: () =>
      deck('Classroom Lesson', 'paper', [
        { layout: 'title', title: 'Lesson Title', body: ['Subject · Grade · Date'] },
        {
          layout: 'titleContent',
          title: 'Learning Objectives',
          body: [lines('By the end of this lesson you will be able to…', 'Objective one', 'Objective two')],
        },
        {
          layout: 'twoContent',
          title: 'Key Vocabulary',
          body: [lines('Term: definition', 'Term: definition'), lines('Term: definition', 'Term: definition')],
        },
        { layout: 'section', title: 'Activity', body: ['Instructions and time limit'] },
        { layout: 'titleContent', title: 'Recap', body: [lines('What we learned today', 'Questions to think about')] },
        { layout: 'titleContent', title: 'Homework', body: [lines('Task and due date', 'Resources')] },
      ]),
  },
  {
    id: 'workshop',
    name: 'Training Workshop',
    description: 'Run a hands-on session with an agenda, modules, exercises and takeaways.',
    kind: 'template',
    build: () =>
      deck('Training Workshop', 'forest', [
        { layout: 'title', title: 'Workshop Title', body: ['Presenter · Date'] },
        {
          layout: 'titleContent',
          title: 'Agenda',
          body: [lines('Introductions', 'Module 1', 'Hands-on exercise', 'Module 2', 'Wrap-up and Q&A')],
        },
        { layout: 'section', title: 'Module 1', body: ['Topic name'] },
        {
          layout: 'titleContent',
          title: 'Exercise',
          body: [lines('What to do', 'Time: 15 minutes', 'What to bring back to the group')],
        },
        {
          layout: 'titleContent',
          title: 'Key Takeaways',
          body: [lines('Takeaway one', 'Takeaway two', 'Where to learn more')],
        },
      ]),
  },
  {
    id: 'portfolio',
    name: 'Personal Portfolio',
    description: 'Showcase who you are, your best work and how to reach you.',
    kind: 'template',
    build: () =>
      deck('Personal Portfolio', 'berry', [
        { layout: 'title', title: 'Your Name', body: ['Designer · Developer · Maker'] },
        {
          layout: 'titleContent',
          title: 'About Me',
          body: [lines('A short introduction', 'What you care about', 'What you’re looking for')],
        },
        {
          layout: 'titleOnly',
          title: 'By the Numbers',
          stats: [
            ['8', 'Years of experience'],
            ['40+', 'Projects shipped'],
            ['12', 'Happy clients'],
          ],
        },
        {
          layout: 'twoContent',
          title: 'Selected Projects',
          body: [lines('Project one', 'Your role and the result'), lines('Project two', 'Your role and the result')],
        },
        { layout: 'section', title: 'Let’s Talk', body: ['email@example.com · yoursite.com'] },
      ]),
  },
  {
    id: 'techtalk',
    name: 'Tech Talk',
    description: 'A developer-style deck for conference and meetup talks.',
    kind: 'template',
    build: () =>
      deck(
        'Tech Talk',
        'terminal',
        [
          { layout: 'title', title: 'talk_title()', body: ['Speaker · Conference 2026'] },
          { layout: 'titleContent', title: '$ whoami', body: [lines('Name and role', 'What you work on', '@handle')] },
          {
            layout: 'twoContent',
            title: 'Before / After',
            body: [lines('Before', 'The slow, painful way'), lines('After', 'The approach in this talk')],
          },
          { layout: 'section', title: 'Demo', body: ['Let’s see it running'] },
          { layout: 'titleContent', title: 'Takeaways', body: [lines('Lesson one', 'Lesson two', 'Link to code and slides')] },
          { layout: 'section', title: 'Thanks!', body: ['Questions?'] },
        ],
        'push',
      ),
  },
  ...ARABIC_TEMPLATES,
  // One blank deck per design theme, like PowerPoint's theme gallery.
  ...THEMES.map<Template>((theme) => ({
    id: `theme-${theme.id}`,
    name: theme.name,
    get description() {
      return t('A blank presentation using the {name} theme. {font} headings.', { name: t(theme.name), font: theme.titleFont });
    },
    kind: 'theme',
    build: () => ({ version: DOC_VERSION, title: t('Untitled presentation'), themeId: theme.id, slides: [makeSlide(theme, 'title')] }),
  })),
];
