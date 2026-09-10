const docx = require('docx');
const fs = require('fs');

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, HeadingLevel, ShadingType,
  PageBreak
} = docx;

const FONT = 'Georgia';
const FONT_CN = 'SimSun';
const TITLE_SIZE = 36;
const BODY_SIZE = 22;
const SMALL_SIZE = 20;
const DATA_SIZE = 18;

function dataBar(values) {
  const cellWidth = Math.floor(13600 / values.length);
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: values.map(() => cellWidth),
    rows: [
      new TableRow({
        children: values.map(v => new TableCell({
          width: { size: cellWidth, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: '1a1a2e' },
          margins: { top: 60, bottom: 60, left: 80, right: 80 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({
              text: v,
              font: FONT,
              size: DATA_SIZE,
              color: 'e0c97f',
              bold: true
            })]
          })]
        }))
      })
    ],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: '3a3a5c' },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: '3a3a5c' },
      left: { style: BorderStyle.SINGLE, size: 1, color: '3a3a5c' },
      right: { style: BorderStyle.SINGLE, size: 1, color: '3a3a5c' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: '3a3a5c' },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '3a3a5c' },
    }
  });
}

function sceneHeader(week, scene, character, stage) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: [3400, 3400, 3400, 3400],
    rows: [
      new TableRow({
        children: [
          ['第 ' + week + ' 周', '场景：' + scene, '男主：' + character, '关系：' + stage].map(t =>
            new TableCell({
              width: { size: 3400, type: WidthType.DXA },
              shading: { type: ShadingType.CLEAR, fill: '0d0d1a' },
              margins: { top: 50, bottom: 50, left: 80, right: 80 },
              children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({
                  text: t,
                  font: FONT,
                  size: DATA_SIZE,
                  color: 'c9b06b',
                  bold: true
                })]
              })]
            })
          )
        ]
      })
    ],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: 'c9b06b' },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: 'c9b06b' },
      left: { style: BorderStyle.SINGLE, size: 1, color: '3a3a5c' },
      right: { style: BorderStyle.SINGLE, size: 1, color: '3a3a5c' },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '3a3a5c' },
      insideHorizontal: { style: BorderStyle.NONE },
    }
  });
}

function bodyP(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 160, line: 360 },
    children: [new TextRun({
      text,
      font: FONT,
      size: opts.size || BODY_SIZE,
      color: opts.color || '2d2d2d',
      bold: opts.bold || false,
      italics: opts.italics || false
    })]
  });
}

function dialogP(speaker, en, cn) {
  const children = [];
  if (speaker) {
    children.push(new TextRun({ text: speaker + '  ', font: FONT, size: BODY_SIZE, bold: true, color: '1a1a2e' }));
  }
  children.push(new TextRun({ text: '"' + en + '"', font: FONT, size: BODY_SIZE, color: '2d2d2d', italics: false }));
  const paras = [new Paragraph({ spacing: { after: 60, line: 360 }, children })];
  if (cn) {
    paras.push(new Paragraph({
      spacing: { after: 160, line: 320 },
      children: [new TextRun({ text: '（' + cn + '）', font: FONT, size: SMALL_SIZE, color: '888888', italics: true })]
    }));
  }
  return paras;
}

function sectionTitle(text) {
  return new Paragraph({
    spacing: { before: 400, after: 200 },
    children: [new TextRun({
      text,
      font: FONT,
      size: 28,
      bold: true,
      color: '1a1a2e'
    })]
  });
}

function choiceBlock(choices) {
  const items = [];
  items.push(new Paragraph({
    spacing: { before: 300, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: 'c9b06b' } },
    children: [new TextRun({ text: '— YOUR MOVE —', font: FONT, size: 24, bold: true, color: 'c9b06b' })]
  }));
  choices.forEach(c => {
    items.push(new Paragraph({
      spacing: { after: 80, line: 340 },
      indent: { left: 360 },
      children: [
        new TextRun({ text: c.label + '  ', font: FONT, size: BODY_SIZE, bold: true, color: '1a1a2e' }),
        new TextRun({ text: c.text, font: FONT, size: BODY_SIZE, color: '2d2d2d' }),
      ]
    }));
    if (c.cn) {
      items.push(new Paragraph({
        spacing: { after: 100, line: 300 },
        indent: { left: 720 },
        children: [new TextRun({ text: '（' + c.cn + '）', font: FONT, size: SMALL_SIZE, color: '888888', italics: true })]
      }));
    }
  });
  items.push(new Paragraph({
    spacing: { after: 80, line: 340 },
    indent: { left: 360 },
    children: [
      new TextRun({ text: 'D.  ', font: FONT, size: BODY_SIZE, bold: true, color: '1a1a2e' }),
      new TextRun({ text: 'Free action — type your own response.', font: FONT, size: BODY_SIZE, color: '2d2d2d' }),
    ]
  }));
  items.push(new Paragraph({
    spacing: { after: 60 },
    indent: { left: 720 },
    children: [new TextRun({ text: '（自由行动——输入你的回应）', font: FONT, size: SMALL_SIZE, color: '888888', italics: true })]
  }));
  return items;
}

function emptyLine() {
  return new Paragraph({ spacing: { after: 100 }, children: [] });
}

// ============================================================
// BUILD DOCUMENT
// ============================================================

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1200, bottom: 1200, left: 1440, right: 1440 }
      }
    },
    children: [

      // ========== TITLE ==========
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: 'COACHELLA 2014', font: FONT, size: 52, bold: true, color: '1a1a2e' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 40 },
        children: [new TextRun({ text: 'Desert Encounters', font: FONT, size: 30, color: 'c9b06b', italics: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [new TextRun({ text: 'Sukey Shen  x  Justin Bieber', font: FONT, size: 24, color: '666666' })]
      }),

      // ========== DATA BAR ==========
      dataBar([
        'Affinity 好感度：5',
        'Mood 心情值：85',
        'Wealth 金钱：999,999',
        'Secrecy 保密度：100',
        'Heat 媒体热度：0',
        'His Stress 压力值：72'
      ]),
      emptyLine(),

      // ========== SCENE HEADER ==========
      sceneHeader('1', 'Coachella Valley · Indio, California', 'Justin Bieber', '陌生人 Strangers'),
      emptyLine(),

      // ========== CONTEXT ==========
      new Paragraph({
        spacing: { after: 100 },
        children: [new TextRun({
          text: 'April 11, 2014 — Friday, Week 1 of Coachella',
          font: FONT, size: SMALL_SIZE, color: '999999', italics: true
        })]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun({
          text: 'Background: Justin Bieber, 20 years old. In the middle of his post-Believe era — the paparazzi won\'t leave him alone, TMZ runs a new headline about him every 48 hours, and he\'s fresh off a DUI arrest in Miami two months ago. He and Selena Gomez are in one of their messy off-again phases. He\'s been spotted partying in LA, drag-racing, egging his neighbor\'s house. The whole world has decided he\'s a lost cause. But underneath the tabloid circus, there\'s still a kid from Stratford, Ontario, who taught himself to play drums, guitar, piano, and trumpet before he turned twelve. His voice — that warm, elastic tenor that can slip from falsetto to a raspy lower register mid-phrase — hasn\'t gone anywhere. If anything, it\'s gotten deeper, richer, tinged with something the Believe tour crowds don\'t fully hear yet.',
          font: FONT, size: SMALL_SIZE, color: '777777'
        })]
      }),

      // ========== BODY ==========

      bodyP('Coachella的第一个夜晚，沙漠的热气还没有完全散去。'),
      emptyLine(),

      bodyP('四月的Indio白天能到三十五度，但太阳落山之后温度会猛降。你穿了一件白色的Zimmermann蕾丝crop top，高腰的浅色牛仔短裤，脚上是一双Isabel Marant的流苏靴。头发没怎么打理，就是自然地散着——黑色的长发在沙漠的干燥空气里反而显得特别顺滑，像是某种跟周围的尘土和混乱格格不入的东西。'),
      emptyLine(),

      bodyP('你其实不太想来。'),
      emptyLine(),

      bodyP('是你UCLA的朋友们硬拉你来的。她们搞到了VIP wristband，说什么"你整天待在Bel Air那个大房子里干嘛，出来玩啊"。你懒得解释你待在家里是因为你家的私人影院刚装了新的IMAX系统，你正在补《纸牌屋》第二季。'),
      emptyLine(),

      bodyP('但你还是来了。'),
      emptyLine(),

      bodyP('此刻你正站在Sahara Tent外面，手里拿着一杯Heineken——你其实不太喜欢啤酒，但VIP区的调酒师做的Margarita实在太甜了。你的两个朋友已经挤进人群去看Disclosure的set了，你一个人靠在围栏边上，百无聊赖地刷手机。'),
      emptyLine(),

      bodyP('你周围的人不停地在偷看你。'),
      emptyLine(),

      bodyP('这不是什么新鲜事。从你十四五岁开始就一直这样。但Coachella这种地方——到处都是穿着花冠和比基尼上衣的model，到处都是Instagram上有几十万粉丝的LA女孩——你还是能感觉到视线集中在你身上。'),
      emptyLine(),

      bodyP('一个戴着Native American羽毛头饰的白人男生端着两杯酒走过来，笑得很灿烂。'),
      emptyLine(),

      ...dialogP('The guy', "Hey, are you here alone? I got an extra drink if you want it.", '嘿，你一个人吗？我多拿了一杯，要不要？'),
      emptyLine(),

      bodyP('他的眼神从你的脸移到你的锁骨，再移到你的腰线，完全没有任何掩饰。'),
      emptyLine(),

      bodyP('你连看都没看他，把手里的Heineken换到另一只手，继续刷手机。'),
      emptyLine(),

      bodyP('他还想说什么，但你旁边突然传来一阵骚动。'),
      emptyLine(),

      bodyP('不是那种Coachella正常的骚动——不是有人醉了或者有人在mosh pit里摔了。是那种特定的、有方向性的骚动。有人在拍照。有人在尖叫。有人在低声说一个名字。'),
      emptyLine(),

      ...dialogP('Crowd whisper', "Is that Bieber? Holy shit, that's Bieber.", '那是Bieber吗？我靠，那是Bieber。'),
      emptyLine(),

      bodyP('你抬起头。'),
      emptyLine(),

      bodyP('一小群人正从Artist Area的方向走过来。中间是一个不算很高的男生——大概五尺九左右——但他的存在感完全不像他的身高。他穿了一件白色的oversized tank top，领口很低，露出锁骨和胸口上方的一小片纹身。深色的drop-crotch pants，脚上是一双红色的Supra。他的头发是那个时期标志性的样子——侧面剃短，上面的金棕色头发往后梳，有点乱，像是故意的又像是懒得打理。'),
      emptyLine(),

      bodyP('他旁边跟着两三个人——一个体型很大的黑人保镖，一个穿格子衫的白人（后来你会知道那是Ryan Good，他当时的stylist兼好朋友），还有一个拿着手机一直在拍的人。'),
      emptyLine(),

      bodyP('Justin Bieber。二十岁。'),
      emptyLine(),

      bodyP('你对他的了解基本来自于你根本没法避免的那些新闻。DUI。跟Selena分分合合。被邻居起诉。在飞机上跟空乘吵架。TMZ每天的头条。你的朋友圈里有人觉得他是joke，有人觉得他可怜，你自己——你其实没怎么想过。'),
      emptyLine(),

      bodyP('但现在他从你面前走过，距离不到三米。'),
      emptyLine(),

      bodyP('你注意到一件事：他的眼睛。'),
      emptyLine(),

      bodyP('电视和杂志上看不出来的那种——他的眼睛是灰蓝色的，在Sahara Tent边上那些彩色灯光的映照下，显得特别浅，特别透。但眼神本身不是那种二十岁男孩该有的。里面有一种你说不上来的东西。不是疲惫，不完全是。更像是一个习惯了被所有人看的人，但已经不太在意被看到什么。'),
      emptyLine(),

      bodyP('他的下巴线条很利落，婴儿肥已经完全褪去了。嘴唇很饱满，上唇微微翘起，给他一种不管什么表情都像是带着一点点不屑的样子。但那不是真的不屑——你看得出来，因为你自己也经常被人说"看起来很凶"。那只是一张天生就长成那样的脸。'),
      emptyLine(),

      bodyP('他路过你的时候，刚才那个端着酒想搭讪你的男生还站在你旁边。Justin的视线扫过来——非常快，像是那种在人群里本能地扫描周围环境的动作。他的目光掠过那个搭讪男生，然后落在你身上。'),
      emptyLine(),

      bodyP('停了大概不到一秒钟。'),
      emptyLine(),

      bodyP('但你注意到了。'),
      emptyLine(),

      bodyP('因为他的表情变了一点。非常细微——嘴角的弧度松了一点，眉心的那道纹路浅了一点。不是smile，不是那种"哇她很好看"的反应。更像是……注意到了。'),
      emptyLine(),

      bodyP('然后他就走过去了。'),
      emptyLine(),

      bodyP('Ryan Good倒是多看了你一眼，然后跟Justin说了句什么。你没听清，音乐太吵了。'),
      emptyLine(),

      bodyP('你转回头继续喝你的Heineken，那个搭讪男生终于识趣地走了。'),
      emptyLine(),

      bodyP('大概过了二十分钟，Disclosure的set还在继续，Sahara Tent里面的bass震得你脚底发麻。你的朋友发来了一条短信：'),
      emptyLine(),

      ...dialogP('Text from Mia', "OMG Sukey where are you?? Kylie Jenner is in the VIP area!! Come!!!", '天哪Sukey你在哪？？Kylie Jenner在VIP区！！快来！！！'),
      emptyLine(),

      bodyP('你翻了个白眼。你对Kylie Jenner没有任何兴趣。'),
      emptyLine(),

      bodyP('你决定去找个安静一点的地方。你沿着Sahara Tent外面的步道往后走，穿过一片装置艺术区——那些巨大的、会发光的蘑菇形状的雕塑和悬挂的LED球。这个区域人比较少，大部分人都挤在各个stage前面。你在一个巨大的金属仙人掌雕塑旁边的草坪上坐下来，把靴子旁边的沙子拍了拍。'),
      emptyLine(),

      bodyP('沙漠的夜空很干净。没有LA那种永远灰蒙蒙的光污染。你能看到星星。'),
      emptyLine(),

      bodyP('你正在享受这难得的安静，突然闻到了一股烟味。不是普通的烟——是那种加州到处都有的、甜甜的、混着草本味道的烟。'),
      emptyLine(),

      bodyP('你转过头。'),
      emptyLine(),

      bodyP('Justin Bieber坐在离你大概五六米远的另一个雕塑底座上。'),
      emptyLine(),

      bodyP('他一个人。保镖站在更远的地方，背对着他，像是在给他一点空间。Ryan Good不见了。那个拿手机拍的人也不在。'),
      emptyLine(),

      bodyP('他坐在那里，一条腿屈起来，手臂搭在膝盖上，手指间夹着一根joint。他在看天。'),
      emptyLine(),

      bodyP('没有了刚才那群人围着的时候的气场，他看起来——小了一号。不是说身高，是整个人的那种存在方式。他看起来就像一个普通的二十岁男孩，在音乐节的角落里抽烟看星星。'),
      emptyLine(),

      bodyP('他也注意到你了。'),
      emptyLine(),

      bodyP('他转过头，看了你一眼。这次不是人群里那种一秒钟的扫视——他真的看了你。'),
      emptyLine(),

      bodyP('沙漠里的风把你的头发吹到脸上。你随手把它拨到耳后。'),
      emptyLine(),

      bodyP('他就在那个瞬间，嘴角微微弯了一下。不是对着镜头的那种笑。是那种——嘴角动了但眼睛没笑的，带着一点curiosity的表情。'),
      emptyLine(),

      bodyP('然后他开口了。声音不大，但在装置艺术区的相对安静里，你听得很清楚。他的声音比你在电视上听到的更粗一点，带着一种慵懒的、低沉的质感——不是刻意压低的那种低，是自然的、二十岁男生变声期彻底结束之后的那种settling down。'),
      emptyLine(),

      ...dialogP('Justin', "You're the only person at Coachella who looks like they don't wanna be here.", '你是Coachella唯一一个看起来不想来的人。'),
      emptyLine(),

      bodyP('他说完这句话，又吸了一口joint，把烟慢慢吐出来。烟雾在夜风里散得很快。'),
      emptyLine(),

      bodyP('他的灰蓝色眼睛透过那一点点残余的烟雾看着你。没有攻击性，没有搭讪的意思，甚至没有什么期待你回应的意思。就是一句observation。'),
      emptyLine(),

      bodyP('像是他看到了一个跟他一样不太属于这里的人，然后忍不住说了出来。'),
      emptyLine(),

      // ========== YOUR MOVE ==========
      ...choiceBlock([
        {
          label: 'A.',
          text: '"Takes one to know one." Then turn back to the sky.',
          cn: '"同类才认得出同类。"然后转回去继续看天空。'
        },
        {
          label: 'B.',
          text: '"I\'m here. That counts for something." Take a sip of your beer and look at him sideways.',
          cn: '"我人在这里，这就够了。"喝一口啤酒，侧眼看他。'
        },
        {
          label: 'C.',
          text: 'Say nothing. But don\'t leave. Just stay sitting there, five meters apart, sharing the same patch of desert sky.',
          cn: '什么都不说。但也不走。就这么坐着，相隔五六米，共享同一片沙漠星空。'
        },
      ]),

      emptyLine(),
      emptyLine(),

      // ========== END DATA BAR ==========
      new Paragraph({
        spacing: { before: 200, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: 'c9b06b' } },
        children: [new TextRun({ text: '— STATUS —', font: FONT, size: 20, bold: true, color: 'c9b06b' })]
      }),
      dataBar([
        'Affinity 好感度：5 (+0)',
        'Mood 心情值：85',
        'Wealth 金钱：999,999',
        'Secrecy 保密度：100',
        'Heat 媒体热度：0',
        'His Stress 压力值：72'
      ]),

      emptyLine(),

      // ========== NOTES ==========
      new Paragraph({
        spacing: { before: 300, after: 120 },
        children: [new TextRun({
          text: 'DM Notes (hidden from player):',
          font: FONT, size: SMALL_SIZE, bold: true, color: 'aaaaaa'
        })]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [new TextRun({
          text: 'Timeline context — April 2014: Justin is between the Believe Tour (ended Dec 2013) and what will become the Journals era. He and Selena broke up again in late January 2014 after a brief reconciliation. His DUI arrest in Miami Beach was January 23. He was charged with egging his neighbor\'s house March 2014. He\'s in a chaotic period — partying with Lil Za and Lil Twist, getting papped everywhere, losing public goodwill fast. But musically, he\'s quietly working — he\'ll release "Confident" ft. Chance the Rapper later this year. The voice is maturing. The kid is not all right, but the talent is still there.',
          font: FONT, size: 18, color: 'bbbbbb', italics: true
        })]
      }),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/home/user/demo1/Coachella_2014_Chapter1.docx', buffer);
  console.log('Done: Coachella_2014_Chapter1.docx');
});
