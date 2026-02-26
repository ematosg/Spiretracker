/*
 * Spire Campaign Manager
 *
 * This single JavaScript file implements the core application logic for
 * managing tabletop RPG campaigns for Spire. The app is designed to run
 * entirely on the client with no backend and persists data via
 * localStorage. Entities (PCs, NPCs, and Organisations) are stored in
 * campaigns, each with a set of relationships forming a conspiracy web.
 * A force-directed graph is rendered with Cytoscape.js to visualise
 * relationships. GM and Player modes control the visibility of secret
 * information. Export/import functionality supports JSON backups and
 * PNG snapshots of the graph.
 */

(() => {
  // Schema version to facilitate future migrations
  const SCHEMA_VERSION = 1;

  // Default relationship types. Users can add custom types in settings.
  const DEFAULT_REL_TYPES = [
    'Ally',
    'Enemy',
    'Rival',
    'Patron/Client',
    'Employer/Employee',
    'Member Of',
    'Family',
    'Romantic',
    'Owes / Debt',
    'Blackmail',
    'Informant',
    'Handler',
    'Targeting',
    'Surveillance',
    'Unknown/Unclear'
  ];

  // Simple random name lists for the name generator
  const DROW_FIRST_NAMES = [
    'Vieriniss', 'Faervel', 'Drisinil', 'Xullrae', 'Sszindy', 'Zesstra',
    'Vesztiira', 'Sharal', 'Belwar', 'Elvraema', 'Malagh', 'Riszar'
  ];
  const DROW_LAST_NAMES = [
    'Xorlarrin', 'Baenre', 'Mizzrym', 'Melarn', 'Helviiryn',
    'Kilsek', 'Despana', 'Kenafin', 'Ousstyl', 'Tsabrin'
  ];
  const MODERN_FIRST_NAMES = [
    'Alex', 'Jamie', 'Morgan', 'Taylor', 'Jordan', 'Riley', 'Casey', 'Reese'
  ];
  const MODERN_LAST_NAMES = [
    'Smith', 'Johnson', 'Taylor', 'Brown', 'Jones', 'Garcia', 'Davis', 'Lee'
  ];

  // Constants for selectable skill and domain options. These lists are used
  // when creating and editing skills/domains to enforce valid names.
  const SKILL_OPTIONS = [
    'Compel', 'Deceive', 'Fight', 'Fix', 'Investigate',
    'Pursue', 'Resist', 'Sneak', 'Steal'
  ];
  const DOMAIN_OPTIONS = [
    'Academia', 'Commerce', 'Crime', 'High Society', 'Low society',
    'Occult', 'Order', 'Religion', 'Technology'
  ];

  // Lightweight in-app fallout guidance by resistance/severity so tables
  // don't need to leave the app mid-session.
  const FALLOUT_GUIDANCE = {
    Blood: {
      Minor: ['Bruised ribs', 'Cut hand', 'Winded and shaky'],
      Moderate: ['Broken finger', 'Concussion symptoms', 'Deep bleeding wound'],
      Severe: ['Crushed limb', 'Internal injuries', 'Near-fatal trauma']
    },
    Mind: {
      Minor: ['Sleepless and rattled', 'Distracted by fear', 'Intrusive memory'],
      Moderate: ['Panic response', 'Paranoia spike', 'Loss of composure'],
      Severe: ['Psychic collapse', 'Dissociative break', 'Severe phobia trigger']
    },
    Silver: {
      Minor: ['Damaged reputation', 'Unexpected expense', 'Social snub'],
      Moderate: ['Debt pressure', 'Public scandal', 'Frozen assets'],
      Severe: ['Financial ruin', 'Total disgrace in high society', 'Powerful creditor vendetta']
    },
    Shadow: {
      Minor: ['Spotted by a watcher', 'Compromised route', 'Rumors of your movements'],
      Moderate: ['Known to hostile faction', 'Safehouse burned', 'Hunted in district'],
      Severe: ['Identity exposed', 'Persistent surveillance net', 'No safe ground left']
    },
    Reputation: {
      Minor: ['Trusted contact offended', 'Street-level mistrust', 'Loss of face'],
      Moderate: ['Faction setback', 'Broken alliance', 'Formal censure'],
      Severe: ['Organization fracture', 'Declared enemy by faction', 'Open political collapse']
    }
  };

  const NPC_TEMPLATES = {
    none: { label: 'No template' },
    guard: {
      label: 'Guard',
      role: 'Guard',
      threatLevel: 'Minor',
      disposition: 'Wary',
      wants: 'Keep order and avoid blame',
      fears: 'Unexpected violence',
      leverage: 'Access to patrol routes',
      inventory: [
        { item: 'Truncheon', quantity: 1, type: 'weapon', stress: 'D3 stress', tags: ['Brutal'] },
        { item: 'Uniform and badge', quantity: 1, type: 'armor', resistance: 1, tags: [] }
      ]
    },
    functionary: {
      label: 'Functionary',
      role: 'Ministry Functionary',
      threatLevel: 'Minor',
      disposition: 'Neutral',
      wants: 'Protect status and paperwork',
      fears: 'Scandal reaching superiors',
      leverage: 'Permits, records, access stamps',
      inventory: [
        { item: 'Official seal', quantity: 1, type: 'other', tags: [] },
        { item: 'Locked folio', quantity: 1, type: 'other', tags: ['Sensitive'] }
      ]
    },
    cultist: {
      label: 'Cultist',
      role: 'Cult Member',
      threatLevel: 'Significant',
      disposition: 'Hostile',
      wants: 'Advance the cult agenda',
      fears: 'Exposure and betrayal',
      leverage: 'Secret meeting sites',
      inventory: [
        { item: 'Ritual knife', quantity: 1, type: 'weapon', stress: 'D3 stress', tags: ['Concealable'] },
        { item: 'Occult iconography', quantity: 1, type: 'other', tags: ['Occult'] }
      ]
    },
    fixer: {
      label: 'Fixer',
      role: 'Street Fixer',
      threatLevel: 'Significant',
      disposition: 'Neutral',
      wants: 'Profit and influence',
      fears: 'Losing network credibility',
      leverage: 'Contacts in multiple districts',
      inventory: [
        { item: 'Hidden pistol', quantity: 1, type: 'weapon', stress: 'D6 stress', tags: ['Concealable', 'Loud'] },
        { item: 'Ledger of favors', quantity: 1, type: 'other', tags: ['Blackmail'] }
      ]
    }
  };

  const SCENE_COMPLICATIONS = [
    'Ministry patrols are unexpectedly present.',
    'A trusted contact is compromised.',
    'Violence erupts nearby with unclear cause.',
    'A key route is blocked or surveilled.',
    'A rumor spreads before the crew arrives.',
    'A faction demands immediate payment or proof.'
  ];
  const SCENE_FACTION_REACTIONS = [
    'The Ministry marks this district as sensitive.',
    'A rival cell starts moving against the party.',
    'A local gang offers help for a price.',
    'A religious faction calls in old obligations.',
    'An organization goes quiet and stops answering.',
    'An enemy openly escalates instead of hiding.'
  ];
  const SCENE_TWISTS = [
    'An old ally appears with conflicting goals.',
    'A mission target is connected to a PC bond.',
    'Collateral damage risk becomes immediate.',
    'Someone the party trusts is feeding information out.',
    'The objective is moved or replaced at the last minute.',
    'A seeming victory creates a worse follow-up problem.'
  ];

  function baseNpcTemplates() {
    return JSON.parse(JSON.stringify(NPC_TEMPLATES));
  }

  function normalizeScenarioPack(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const id = String(raw.id || '').trim() || generateId('pack');
    const name = String(raw.name || '').trim() || 'Scenario Pack';
    const source = String(raw.source || 'import').trim() || 'import';
    const enabled = raw.enabled !== false;
    const prompts = raw.scenePrompts || {};
    const npcTemplates = (raw.npcTemplates && typeof raw.npcTemplates === 'object') ? raw.npcTemplates : {};
    const clean = {
      id,
      cloudCampaignId: '',
      name,
      source,
      enabled,
      scenePrompts: {
        complications: Array.isArray(prompts.complications) ? prompts.complications.map(String).map((s) => s.trim()).filter(Boolean) : [],
        factionReactions: Array.isArray(prompts.factionReactions) ? prompts.factionReactions.map(String).map((s) => s.trim()).filter(Boolean) : [],
        twists: Array.isArray(prompts.twists) ? prompts.twists.map(String).map((s) => s.trim()).filter(Boolean) : []
      },
      npcTemplates: {}
    };
    Object.entries(npcTemplates).forEach(([key, value]) => {
      if (!key || key === 'none' || !value || typeof value !== 'object') return;
      clean.npcTemplates[String(key)] = JSON.parse(JSON.stringify(value));
    });
    if (
      !clean.scenePrompts.complications.length &&
      !clean.scenePrompts.factionReactions.length &&
      !clean.scenePrompts.twists.length &&
      !Object.keys(clean.npcTemplates).length
    ) return null;
    return clean;
  }

  function getActiveScenarioPacks(camp = currentCampaign()) {
    const packs = Array.isArray(camp && camp.scenarioPacks) ? camp.scenarioPacks : [];
    return packs.filter((p) => p && p.enabled !== false);
  }

  function getEffectiveNpcTemplates(camp = currentCampaign()) {
    const merged = baseNpcTemplates();
    getActiveScenarioPacks(camp).forEach((pack) => {
      Object.entries(pack.npcTemplates || {}).forEach(([key, tpl]) => {
        if (!key || !tpl || typeof tpl !== 'object') return;
        merged[key] = JSON.parse(JSON.stringify(tpl));
      });
    });
    return merged;
  }

  function getEffectiveScenePromptPools(camp = currentCampaign()) {
    const pools = {
      complications: SCENE_COMPLICATIONS.slice(),
      factionReactions: SCENE_FACTION_REACTIONS.slice(),
      twists: SCENE_TWISTS.slice()
    };
    getActiveScenarioPacks(camp).forEach((pack) => {
      const prompts = pack.scenePrompts || {};
      if (Array.isArray(prompts.complications)) pools.complications.push(...prompts.complications);
      if (Array.isArray(prompts.factionReactions)) pools.factionReactions.push(...prompts.factionReactions);
      if (Array.isArray(prompts.twists)) pools.twists.push(...prompts.twists);
    });
    pools.complications = Array.from(new Set(pools.complications.filter(Boolean)));
    pools.factionReactions = Array.from(new Set(pools.factionReactions.filter(Boolean)));
    pools.twists = Array.from(new Set(pools.twists.filter(Boolean)));
    return pools;
  }

  function defaultFalloutGuidance() {
    return JSON.parse(JSON.stringify(FALLOUT_GUIDANCE));
  }

  // Durance options for characters. Each option modifies the character
  // by granting skills, domains or resistance bonuses. The format is
  // { skills: [ ... ], domains: [ ... ], resistances: [ {name,value} ] }.
  const DURANCE_OPTIONS = [
    'ACOLYTE', 'AGENT', 'BUILDER', 'DEALER', 'DUELLIST', 'ENLISTED',
    'GUARD', 'HUMAN EMISSARY', 'HUNTER', 'INFORMATION BROKER', 'KILLER',
    'LABOURER', 'OCCULTIST', 'PERSONAL ASSISTANT', 'PET', 'SAGE', 'SPY',
    'Kept a low profile in Derelictus', 'Fought to protect your community',
    'Led a doomed uprising', 'Joined a cult or two', 'Toughed it out in Red Row',
    'Fell in with a gang of thieves', 'Spent your time in jail', 'Hid in plain sight',
    'Helped the Ministry wage their war'
  ];

  const DURANCE_EFFECTS = {
    'ACOLYTE': { resistances: [{ name: 'Mind', value: 2 }], domains: ['Religion'] },
    'AGENT': { resistances: [{ name: 'Shadow', value: 2 }], domains: ['Crime'] },
    'BUILDER': { skills: ['Fix'], domains: ['Technology'] },
    'DEALER': { skills: ['Compel'], domains: ['Commerce'] },
    'DUELLIST': { skills: ['Fight'], domains: ['High Society'] },
    'ENLISTED': { resistances: [{ name: 'Blood', value: 2 }], skills: ['Fight'] },
    'GUARD': { resistances: [{ name: 'Reputation', value: 2 }], domains: ['Order'] },
    'HUMAN EMISSARY': { domains: ['Technology', 'Commerce'] },
    'HUNTER': { skills: ['Pursue', 'Sneak'] },
    'INFORMATION BROKER': { resistances: [{ name: 'Shadow', value: 2 }], skills: ['Investigate'] },
    'KILLER': { skills: ['Sneak', 'Fight'] },
    'LABOURER': { resistances: [{ name: 'Blood', value: 2 }] },
    'OCCULTIST': { resistances: [{ name: 'Shadow', value: 2 }], domains: ['Occult'] },
    'PERSONAL ASSISTANT': { resistances: [{ name: 'Silver', value: 2 }], skills: ['Compel'] },
    'PET': { resistances: [{ name: 'Silver', value: 2 }], domains: ['High Society'] },
    'SAGE': { resistances: [{ name: 'Mind', value: 2 }], domains: ['Academia'] },
    'SPY': { skills: ['Sneak', 'Deceive'] },
    'Kept a low profile in Derelictus': { resistances: [{ name: 'Shadow', value: 2 }], domains: ['Low society'] },
    'Fought to protect your community': { resistances: [{ name: 'Reputation', value: 2 }], domains: ['Crime'] },
    'Led a doomed uprising': { skills: ['Compel'], domains: ['Low society'] },
    'Joined a cult or two': { domains: ['Religion', 'Occult'] },
    'Toughed it out in Red Row': { skills: ['Fight'], domains: ['Crime'] },
    'Fell in with a gang of thieves': { skills: ['Sneak', 'Steal'] },
    'Spent your time in jail': { resistances: [{ name: 'Blood', value: 2 }], domains: ['Crime'] },
    'Hid in plain sight': { skills: ['Deceive'], domains: ['High society'] },
    'Helped the Ministry wage their war': { resistances: [{ name: 'Shadow', value: 2 }], skills: ['Resist'] }
  };

  /**
   * Class effects for PCs. Each class defines resistances bonuses, skills,
   * domains, bond prompts, inventory options, core abilities, and
   * advances at low, medium and high tiers. Inventory options are
   * represented as an array of option groups; each group contains
   * a label and a list of items to be added to the character's
   * inventory when selected. Core abilities and advances are lists
   * of strings. Resistances bonuses use the same structure as
   * DURANCE_EFFECTS (name/value pairs).
   */
  const CLASS_EFFECTS = {
    'Azurite': {
      resistances: [{ name: 'Silver', value: 2 }, { name: 'Reputation', value: 2 }],
      skills: ['Compel', 'Deceive'],
      domains: ['Commerce', 'High Society', 'Low society'],
      refresh: 'Carry out a deal that benefits you more than it does the other party.',
      bondPrompts: [
        'You have an individual-level bond with someone who buys, sells, or smuggles things for a living. Name them and what they’re most interested in.',
        'You have a bond with one of the other PCs whom you helped out of debt. Say who, and why they got into debt in the first place.'
      ],
      inventoryOptions: [
        {
          label: 'Standard kit',
          items: [
            { item: 'Blue robes, many layers', quantity: 1, type: 'armor', resistance: 0, tags: [] },
            { item: 'Gold jewellery (coins from overseas)', quantity: 1, type: 'other', tags: [] },
            { item: 'Buckler of Azur', quantity: 1, type: 'armor', resistance: 1, tags: ['holy symbol'] },
            { item: 'Serious-looking club', quantity: 1, type: 'weapon', stress: 'D3 stress', tags: ['Brutal'] }
          ]
        },
        {
          label: 'Opulent kit',
          items: [
            { item: 'Three sets of beautiful robes and girdles, each in slightly different shades of blue', quantity: 1, type: 'armor', resistance: 0, tags: [] },
            { item: 'Golden necklaces, nose-rings, and bracelets bearing the symbol of Azur', quantity: 1, type: 'other', tags: [] },
            { item: 'Bodyguard (Weapon)', quantity: 1, type: 'weapon', stress: 'D6 stress', tags: ['Tiring'] }
          ],
          note: 'If you choose the second option, you’re protected by an agent under your employ. Name and describe your bodyguard and note down two things they hate.'
        }
      ],
      coreAbilities: ['CUT A DEAL', 'HEART’S DESIRE'],
      advances: {
        low: ['GOLDEN TONGUE', 'IGNOBLE TACTICS', 'GOLD-BLOODED', 'HIDDEN STASHES', 'THE GOLDEN GOD’S ARCANA', 'BUY FRIENDS', 'GLUTTON’S COIN'],
        medium: ['TRUE BLUE', 'DESPERATE BARGAIN', 'GOLDEN QUILL', 'ON THE TOSS OF A COIN', 'AZUR’S GRACE', 'THE GOLDEN GOD’S GUIDANCE', 'BUY LOYALTY'],
        high: ['BUY SOME TIME', 'GOLDEN HANDSHAKE', 'BUY ANYTHING', 'BUY POWER']
      }
    },
    'Blood Witch': {
      resistances: [{ name: 'Blood', value: 3 }, { name: 'Shadow', value: 1 }],
      skills: ['Deceive', 'Resist'],
      domains: ['Occult', 'Low society'],
      refresh: 'Share a moment of intimacy with another person.',
      bondPrompts: [
        'You have captured a creature and fed your diseased blood to it, turning it to your will and enhancing its intelligence. Choose a small common creature such as a cat, toad, crow, snake, spider, or raven, and gain it as an individual-level bond. The creature has a physical tell that indicates it is under the influence of black magic, such as compound eyes, additional legs, strange markings, or horns.',
        'You have tasted the blood of another character, and learned a secret about their past (or future). What did you learn, and how often do you remind them of it?'
      ],
      inventoryOptions: [
        {
          label: 'Blood Witch kit',
          items: [
            { item: 'Athame', quantity: 1, type: 'weapon', stress: 'D3 stress', tags: ['Bloodbound'] },
            { item: 'Several sets of tattered, weird-looking clothing', quantity: 1, type: 'other', tags: [] },
            { item: 'A collection of occult ephemera', quantity: 1, type: 'other', tags: [] }
          ]
        }
      ],
      coreAbilities: ['NIGHT TERROR', 'TRUE FORM'],
      advances: {
        low: ['ARTERIAL SPRAY', 'BLIND EYE CURSE', 'BLOOD-BOUND COMPANION', 'BLOODY MASK', 'BLOOD WARD', 'EVIL EYE'],
        medium: ['CLOSE THE WOUND', 'CORPUS DESAN', 'HEART\'S-BLOOD THRALL', 'LAIR', 'TORRENT', 'MANNEQUIN CURSE', 'WENDING CORRIDORS'],
        high: ['A DARK AND BLASTED LAND', 'CHEVAL', 'UNKILLABLE']
      }
    },
    'Bound': {
      resistances: [{ name: 'Blood', value: 1 }, { name: 'Shadow', value: 2 }],
      skills: ['Fight', 'Sneak', 'Pursue'],
      domains: ['Low society', 'Crime'],
      refresh: 'Bring a criminal to justice.',
      bondPrompts: [
        'You have an individual-level bond with a member of the downtrodden underclass. Name them, and name the thing that’s most important to them.',
        'You have a bond with one of the other PCs who you rescued from a dangerous situation. Describe the situation they found themselves in.'
      ],
      inventoryOptions: [
        {
          label: 'Standard kit',
          items: [
            { item: 'Light leather armour', quantity: 1, type: 'armor', resistance: 2, tags: [] },
            { item: 'Ceremonial red binding ropes and mask', quantity: 1, type: 'other', tags: [] },
            { item: 'Sturdy leather gloves', quantity: 1, type: 'other', tags: [] },
            { item: 'Climbing gear and ropes', quantity: 1, type: 'other', tags: [] }
          ]
        },
        {
          label: 'God-knife',
          items: [ { item: 'God-knife', quantity: 1, type: 'weapon', stress: 'D3 stress', tags: ['Concealable', 'Bound'] } ]
        },
        {
          label: 'God-axe',
          items: [ { item: 'God-axe', quantity: 1, type: 'weapon', stress: 'D6 stress', tags: ['Bound'] } ]
        }
      ],
      coreAbilities: ['SURPRISE INFILTRATION', 'BOUND BLADE'],
      advances: {
        low: ['THE SECRET OF BINDING', 'THE SECRET OF SECOND SKIN', 'THE SECRET OF THE CROWD', 'THE SECRET OF FLIGHT', 'THE SECRET OF LOOSE TONGUES', 'THE SECRET OF FEAR', 'THE SECRET OF LUCKY BREAKS'],
        medium: ['THE SAINT OF BLADES', 'THE SAINT OF BLOOD', 'THE SAINT OF BINDING', 'THE SAINT OF HIDDEN FACES', 'THE SAINT OF WAYS', 'THE SAINT OF LAST STANDS'],
        high: ['THE GOD OF SLAUGHTER', 'THE GOD OF SHADOWS', 'THE GOD OF PERCH', 'THE GOD OF GETTING EVEN']
      }
    },
    'Carrion-Priest': {
      resistances: [{ name: 'Blood', value: 2 }, { name: 'Reputation', value: 2 }],
      skills: ['Pursue', 'Sneak'],
      domains: ['Religion', 'Low society'],
      refresh: 'Complete a hunt and take your quarry.',
      bondPrompts: [
        'You have a street-level bond with the faithful of charnel – a collection of worshippers of the corpse-eater god who live in New Heaven. Name three of them, and what’s weird about them.',
        'You have a bond with another PC who you have helped deal with a death – either by guiding them through the grieving process or disposing of the body. Say who it was, and who died.'
      ],
      inventoryOptions: [
        {
          label: 'Crossbow kit',
          items: [
            { item: 'Leathers and robes', quantity: 1, type: 'armor', resistance: 2, tags: [] },
            { item: 'Hyena', quantity: 1, type: 'other', tags: [] },
            { item: 'Heavy-pull crossbow', quantity: 1, type: 'weapon', stress: 'D8 stress', tags: ['Ranged', 'Reload', 'Unreliable'] },
            { item: 'Knife', quantity: 1, type: 'weapon', stress: 'D3 stress', tags: ['Concealable'] }
          ]
        },
        {
          label: 'War-cleaver kit',
          items: [
            { item: 'Leathers and robes', quantity: 1, type: 'armor', resistance: 2, tags: [] },
            { item: 'Hyena', quantity: 1, type: 'other', tags: [] },
            { item: 'War-cleaver', quantity: 1, type: 'weapon', stress: 'D6 stress', tags: [] },
            { item: 'Preyhook', quantity: 1, type: 'weapon', stress: 'D3 stress', tags: ['Ranged', 'Stunning'] }
          ]
        }
      ],
      coreAbilities: ['HYENA', 'LAY OF THE LAND'],
      advances: {
        low: ['NEW TRICKS', 'CACKLE', 'MURDER OF CROWS', 'RIP AND TEAR', 'DEAD FLESH', 'CHARNEL’S MARK'],
        medium: ['GHOST SPEAKER', 'RED FEAST', 'MASSACRE', 'ALPHA', 'FORM OF THE CORVID', 'RED OF BEAK AND TALON'],
        high: ['BLOODHUNT', 'A FLOCK OF NIGHT-BLACK TERRORS', 'TASTE LIFE', 'FORM OF THE GREAT CARRION-EATER']
      }
    },
    'Firebrand': {
      resistances: [{ name: 'Reputation', value: 3 }, { name: 'Shadow', value: 1 }],
      skills: ['Compel', 'Steal'],
      domains: ['Low society', 'Crime'],
      refresh: 'Take something back from those who would oppress you.',
      bondPrompts: [
        'You have two individual-level bonds with folk who are sympathetic to your goals. Pick two domains, and create an NPC bond for each of them.',
        'You have a bond with one of the other PCs who you recruited to the cause. Say who, and say what it was that tipped them over the edge.'
      ],
      inventoryOptions: [
        {
          label: 'Light weapon kit',
          items: [ { item: 'Knife or sap or brass knuckles', quantity: 1, type: 'weapon', stress: 'D3 stress', tags: ['Concealable'] } ]
        },
        {
          label: 'Heavy tool kit',
          items: [ { item: 'Sledgehammer or pickaxe', quantity: 1, type: 'weapon', stress: 'D6 stress', tags: ['Tiring'] } ]
        },
        {
          label: 'Revolver kit',
          items: [ { item: 'Crow-pattern revolver', quantity: 1, type: 'weapon', stress: 'D6 stress', tags: ['Unreliable'] } ]
        },
        {
          label: 'Shotgun kit',
          items: [ { item: 'Buzzard sawn-off', quantity: 1, type: 'weapon', stress: 'D6 stress', tags: ['Reload', 'Point-Blank'] } ]
        }
      ],
      coreAbilities: ['LEAD FROM THE FRONT', 'DRAW A CROWD'],
      advances: {
        low: ['FIGHT THE POWER', 'NOBLE SACRIFICE', 'FORCE OF PERSONALITY', 'ALWAYS OUTNUMBERED, NEVER OUTRUN', 'BROTHERS IN ARMS', 'GODDESS’ CHOSEN'],
        medium: ['ME AND THIS ARMY', 'THE PEOPLE’S CHAMPION', 'SCAPEGOAT', 'MAKE AN EXAMPLE', 'FRIENDS IN LOW PLACES', 'UNTOUCHABLE'],
        high: ['MY NAME IS LEGION', 'THE MEANS OF DESTRUCTION', 'IRON WILL', 'YOU CAN’T KILL AN IDEA']
      }
    },
    'Idol': {
      resistances: [{ name: 'Silver', value: 1 }, { name: 'Mind', value: 1 }, { name: 'Reputation', value: 2 }],
      skills: ['Deceive', 'Compel'],
      domains: ['High society', 'Occult'],
      refresh: 'Someone feels deeply moved when they witness your art.',
      bondPrompts: [
        'You have a street-level bond to your adoring fans. Name three of them, and what the group is most excited to see next.',
        'You have a bond with another PC who you know has feelings for you, even if they wouldn’t admit it. Describe the moment when you knew for definite.'
      ],
      inventoryOptions: [
        {
          label: 'Idol kit',
          items: [
            { item: 'Several sets of flattering clothing', quantity: 1, type: 'other', tags: [] },
            { item: 'Tools to create and perform your chosen art', quantity: 1, type: 'other', tags: [] },
            { item: 'Small gifts and trinkets from your fans', quantity: 1, type: 'other', tags: [] },
            { item: 'Knife', quantity: 1, type: 'weapon', stress: 'D3 stress', tags: ['Concealable'] }
          ]
        }
      ],
      coreAbilities: ['LIFE AND SOUL OF THE PARTY', 'GLAMOUR'],
      advances: {
        low: ['CENTRE OF ATTENTION', 'GRACE', 'WHO ARE THEY?', 'MAJESTY', 'DISHARMONY', 'INSTILL EMOTION', 'INCORRUPTIBLE'],
        medium: ['BEAUTY IS TRUTH', 'UNTOUCHABLE', 'SPITE', 'KILL FOR ME', 'PAINT WITH BLOOD', 'RENDER UNTO ME'],
        high: ['TRUTH IS BEAUTY', 'HAPPY TO HELP', 'SOUL’S PORTRAIT', 'PERFECTION']
      }
    },
    'Knight': {
      resistances: [{ name: 'Blood', value: 1 }, { name: 'Silver', value: 2 }, { name: 'Reputation', value: 1 }],
      skills: ['Fight', 'Compel'],
      domains: ['Low society', 'Crime'],
      refresh: 'Engage in reckless excess.',
      bondPrompts: [
        'You have an individual-level bond with your squire – a young dark elf serving you with an eye to becoming a Knight themselves some day. Name them and say whether they’re idealistic or cynical about the whole affair.',
        'You have a bond with another one of the PCs – you and them used to go drinking, and still do on occasion. Describe the wildest thing you two got up to on one of your legendary nights out.'
      ],
      inventoryOptions: [
        {
          label: 'Knight Quarter-Plate + Greatsword',
          items: [ { item: 'Knight Quarter-Plate', quantity: 1, type: 'armor', resistance: 3, tags: ['Heavy'] }, { item: 'Greatsword', quantity: 1, type: 'weapon', stress: 'D8 stress', tags: ['Tiring'] } ]
        },
        {
          label: 'Knight Quarter-Plate + Sword and Pistol',
          items: [ { item: 'Knight Quarter-Plate', quantity: 1, type: 'armor', resistance: 3, tags: ['Heavy'] }, { item: 'Sword', quantity: 1, type: 'weapon', stress: 'D6 stress', tags: [] }, { item: 'Grackler Pistol', quantity: 1, type: 'weapon', stress: 'D6 stress', tags: ['Brutal', 'Ranged', 'One-shot'] } ]
        },
        {
          label: 'Knight Quarter-Plate + Knightly Lance',
          items: [ { item: 'Knight Quarter-Plate', quantity: 1, type: 'armor', resistance: 3, tags: ['Heavy'] }, { item: 'Knightly Lance', quantity: 1, type: 'weapon', stress: 'D6 stress', tags: ['Piercing', 'Surprising'] } ]
        }
      ],
      coreAbilities: ['PUBCRAWLER', 'PICK A FIGHT', 'LAW OF THE DOCKS'],
      advances: {
        low: ['CAROUSE', 'OUSTER', 'BRAGGADOCIO', 'KNIGHT-ADMIRAL', 'BULWARK', 'KNIGHT-PROTECTOR', 'THE CROWD GOES WILD'],
        medium: ['ARMOUR-KENNING', 'RACONTEUR', 'BRING IT ON', 'DIRTY FIGHTING', 'DO YOU KNOW WHO I AM?', 'RIGHT PLACE, WRONG TIME', 'LAW OF THE LAND'],
        high: ['FORTRESS PLATE', 'PULL THE SWORD FROM THE STONE', 'SLAY THE DRAGON', 'SEEK THE GRAIL']
      }
    },
    'Lahjan': {
      resistances: [{ name: 'Mind', value: 2 }, { name: 'Reputation', value: 2 }],
      skills: ['Fix', 'Resist'],
      domains: ['Religion', 'Low society'],
      refresh: 'Help those who cannot help themselves.',
      bondPrompts: [
        'You have an individual-level bond with an NPC member of the congregation who is sympathetic to your goals. Name them, and what they’re getting out of the relationship.',
        'You have a bond with a PC who you’ve helped overcome sickness, injury or addiction in the past. Say who it was, and what the problem was.'
      ],
      inventoryOptions: [
        {
          label: 'Nansan kit',
          items: [ { item: 'Ceremonial robes', quantity: 1, type: 'other', tags: [] }, { item: 'Jewellery set (wooden and silver)', quantity: 1, type: 'other', tags: [] }, { item: "L'od Nansan Knife", quantity: 1, type: 'weapon', stress: 'D3 stress', tags: ['Concealable'] }, { item: "Healer’s kit", quantity: 1, type: 'other', tags: [] } ]
        },
        {
          label: 'Limyé-Anjhan kit',
          items: [ { item: 'Ceremonial robes', quantity: 1, type: 'other', tags: [] }, { item: 'Jewellery set (wooden and silver)', quantity: 1, type: 'other', tags: [] }, { item: "Moonsilver staff", quantity: 1, type: 'weapon', stress: 'D3 stress', tags: ['Conduit'] } ]
        }
      ],
      coreAbilities: ['RITE OF RESPITE', 'MOONLIGHT'],
      advances: {
        low: ['BUILD BRIDGES', 'BURN BRIDGES', 'BEDSIDE MANNER', 'FRIEND TO THE DOWNTRODDEN', 'OUR LADY’S CALM', 'RITE OF THE SILVER SANCTUARY', 'SCRYATRIX NASCEN'],
        medium: ['SHIMMERING IMAGE', 'RITE OF THE THREE SISTERS', 'SCRYATRIX INANIS', 'OUR LADY’S KISS', 'OUR LADY’S CURSE', 'PERFECT MIRROR'],
        high: ['BODY OF SILVER LIGHT', 'OUR LADY’S MARTYR', 'SCRYATRIX DEMEN', 'BEYOND THE GARDEN GATE']
      }
    },
    'Masked': {
      resistances: [{ name: 'Silver', value: 1 }, { name: 'Mind', value: 1 }, { name: 'Shadow', value: 2 }],
      skills: ['Resist', 'Compel'],
      domains: ['High society', 'Order'],
      refresh: 'Show someone they should not have underestimated you.',
      bondPrompts: [
        'You have one street-level bond with the servants of your old master. Name three of them and describe their jobs, and note down your master’s name and the worst thing they ever did to you or someone else under their power.',
        'You have a bond with another PC who you assisted during their durance. Who was it, and how did you help them out?'
      ],
      inventoryOptions: [
        {
          label: 'Masked kit: Pistol',
          items: [ { item: 'Your Mask', quantity: 1, type: 'other', tags: [] }, { item: 'Nice clothing x2', quantity: 1, type: 'other', tags: [] }, { item: 'Servant Mask', quantity: 1, type: 'other', tags: [] }, { item: 'Hawk Duelling Pistol', quantity: 1, type: 'weapon', stress: 'D6 stress', tags: ['Piercing', 'Ranged', 'One-shot'] } ]
        },
        {
          label: 'Masked kit: Dagger',
          items: [ { item: 'Your Mask', quantity: 1, type: 'other', tags: [] }, { item: 'Nice clothing x2', quantity: 1, type: 'other', tags: [] }, { item: 'Servant Mask', quantity: 1, type: 'other', tags: [] }, { item: 'Dagger', quantity: 1, type: 'weapon', stress: 'D3 stress', tags: ['Concealable'] } ]
        }
      ],
      coreAbilities: ['SMELL STATUS', 'SERVANT TO THE HIGH ONES'],
      advances: {
        low: ['CITIZEN’S MASK', 'INSTITUTIONAL FALSEHOOD', 'INNER MASK OF CALM', 'ONE OF THE STAFF', 'ONE EYE OPEN', 'DRESS FOR SUCCESS'],
        medium: ['MASK OF THE LOVER', 'MASK OF THE KILLER', 'MASK OF PLENTY', 'MOUTHLESS MASK', 'MIRROR-MASK'],
        high: ['THE MASTERLESS MASK', 'GESTALT', 'PANTHEON MASK']
      }
    },
    'Midwife': {
      resistances: [{ name: 'Blood', value: 2 }, { name: 'Reputation', value: 1 }, { name: 'Mind', value: 1 }],
      skills: ['Fix', 'Fight'],
      domains: ['Occult', 'Low society'],
      refresh: 'Defend the defenceless.',
      bondPrompts: [
        'You have a street-level bond with the Order of Midwives, and are an active member. Name your immediate superior, who does not know you work for the Ministry, and one colleague, who does.',
        'You have a bond with another player character, whose life you saved when no one else would. Say who, and what they’d done to ostracise themselves from their community.'
      ],
      inventoryOptions: [
        {
          label: 'Twin Razors',
          items: [ { item: 'Ceremonial silk robes', quantity: 1, type: 'other', tags: [] }, { item: 'Twin Razors', quantity: 1, type: 'weapon', stress: 'D6 stress', tags: ['Concealable', 'Unreliable'] } ]
        },
        {
          label: 'Weighted chain',
          items: [ { item: 'Ceremonial silk robes', quantity: 1, type: 'other', tags: [] }, { item: 'Weighted chain', quantity: 1, type: 'weapon', stress: 'D3 stress', tags: ['Concealable', 'Stunning'] } ]
        }
      ],
      coreAbilities: ['MARTYR', 'PROTECTOR’S EYE'],
      advances: {
        low: ['CANTICLE OF REMAKING', 'WEB OF THE MISTRESS', 'HANDS OF THE MOTHER', 'BLESSING OF ISHKRAH', 'PLUCK THE WEB', 'EYES OF ISHKRAH', 'RITE OF STILLED MIND'],
        medium: ['WEAVE THE WEB', 'CHITINOUS SHELL', 'COCOON OF REBIRTH', 'VENOMOUS MANDIBLES', 'WALL-WALKER', 'ARACHNID BODY', 'SUMMON SWEETLINGS'],
        high: ['ISHKRAH’S PERFECT COCOON', 'PURGE', 'NO MAGIC BUT THE MAGIC OF MY MISTRESS', 'FORM OF ISHKRAH']
      }
    },
    'Vermissian Sage': {
      resistances: [{ name: 'Mind', value: 3 }, { name: 'Shadow', value: 1 }],
      skills: ['Investigate', 'Compel'],
      domains: ['Academia', 'Occult', 'Technology'],
      refresh: 'Uncover hidden information.',
      bondPrompts: [
        'You have an individual-level bond with an academic, researcher or guardian of the Vault. Name them, and their specialty.',
        'You have a bond with another PC – you know a secret about them. Say who it is, what the secret is, and whether they know you know or not.'
      ],
      inventoryOptions: [
        {
          label: 'Crossbow kit',
          items: [ { item: 'Folding crossbow', quantity: 1, type: 'weapon', stress: 'D6 stress', tags: ['Ranged', 'Concealable', 'One-shot'] } ]
        },
        {
          label: 'Dagger kit',
          items: [ { item: 'Dagger', quantity: 1, type: 'weapon', stress: 'D3 stress', tags: ['Concealable'] }, { item: 'Padded vest', quantity: 1, type: 'armor', resistance: 1, tags: [] } ]
        }
      ],
      coreAbilities: ['BACK DOOR', 'THE VAULT', 'OBSESSIVE RESEARCHER'],
      advances: {
        low: ['FIND CONNECTION', 'DEAD DROP', 'MENTAL DIRECTORY', 'THE LOCKED STACKS', 'THORNED TONGUE', 'THAT DIDN’T HAPPEN'],
        medium: ['POCKET GUIDE', 'UNSPEAKABLE', 'MEMORY BLANK', 'YNASTIC MEMORY', 'VERMISSIAN DROP'],
        high: ['UNREADABLE', 'ANASTOMOSIS', 'THE GLASS LIBRARY', 'REWRITE']
      }
    }
  };
  const WEAPON_STRESS_LEVELS = [
    '1 stress', 'D3 stress', 'D6 stress', 'D8 stress'
  ];
  const WEAPON_TAGS = [
    'Accurate', 'Bloodbound', 'Bound', 'Brutal', 'Concealable', 'Conduit',
    'Dangerous', 'Defensive', 'Devastating', 'Double-Barrelled',
    'Extreme Range', 'Masterpiece', 'One-Shot', 'Ongoing Dx', 'Parrying',
    'Piercing', 'Point-Blank', 'Ranged', 'Reload', 'Scarring', 'Spread Dx',
    'Surprising', 'Stunning', 'Tiring', 'Unreliable'
  ];
  const ARMOR_TAGS = [
    'Assault', 'Camouflaged', 'Concealable', 'Implacable'
  ];

  // Random organisation name generator. Combines a prefix with a suffix for
  // evocative faction names. These lists can be expanded or modified to
  // better suit your campaign.
  const ORG_NAME_PREFIXES = [
    'House', 'Cult of', 'Guild of', 'Circle of', 'Syndicate of', 'Order of',
    'Brotherhood of', 'Society of', 'Cabal of', 'Consortium of', 'League of',
    'Institute of', 'Fellowship of'
  ];
  const ORG_NAME_SUFFIXES = [
    'Whispers', 'Shadows', 'Bone', 'Twilight', 'Blood', 'Steel', 'Glass',
    'Secrets', 'Dust', 'Echoes', 'Ash', 'Fate', 'Flux', 'Serpents',
    'Embers', 'Silence', 'Chains', 'Night', 'Knives', 'Mists'
  ];

  // ---------------------------------------------------------------------------
  // Graph implementation using HTML5 Canvas
  //
  // The app originally used Cytoscape.js for graph rendering. Because
  // external CDNs may be unreachable in this environment, a custom
  // canvas-based graph is implemented here. Nodes and edges are drawn
  // manually, supporting pan, zoom, drag, selection and simple layouts.

  // Canvas elements and drawing context. These are initialised in
  // setupGraph().
  let graphCanvas = null;
  let graphCtx = null;
  // Array of node objects currently visible in the graph. Each node has
  // {id, ent, x, y, color, shape}. The x and y coordinates are in
  // arbitrary graph space (not pixels). Graph positions are stored in
  // currentCampaign().positions between sessions.
  let graphNodes = [];
  // Array of edge objects currently visible in the graph. Each edge has
  // {id, source, target, type, directed, sourceKnows, targetKnows, secret}.
  let graphEdges = [];
  let graphHoveredNodeId = null;
  let graphSelectedNodeId = null;
  let graphTooltip = { visible: false, text: '', x: 0, y: 0 };
  // Graph view state including current scale and translation, drag
  // behaviour and tracking of dragging/panning.
  const graphState = {
    scale: 1,
    translateX: 0,
    translateY: 0,
    draggingNodeId: null,
    dragOffsetX: 0,
    dragOffsetY: 0,
    isPanning: false,
    lastX: 0,
    lastY: 0,
    movedDuringDrag: false,
    focusMode: false
  };

  /**
   * Generate a random organisation name. Combines a prefix and a suffix.
   */
  function randomOrgName() {
    return `${randomFrom(ORG_NAME_PREFIXES)} ${randomFrom(ORG_NAME_SUFFIXES)}`;
  }

  // Global state. Contains all campaigns and UI state like current selection.
  const state = {
    campaigns: {},      // Map of campaignId -> campaign
    currentCampaignId: null,
    selectedEntityId: null,
    selectedRelId: null,
    gmMode: true,      // True if GM view, false for player
    darkMode: false,
    relTypes: DEFAULT_REL_TYPES.slice(),
    users: {},
    currentUser: null,
    clientId: generateId('client'),
    realtimeChannel: null,
    supabaseRealtimeChannel: null,
    syncConfig: null,
    realtimeTransport: 'local',
    telemetryConfig: null,
    lastSeenRevision: '',
    syncConflictActive: false,
    localEditsSinceConflict: 0
  };
  let authMode = 'login';
  const logFilterState = {
    session: 'all',
    query: '',
    actor: 'all',
    actionType: 'all',
    page: 1,
    pageSize: 50
  };
  const messageFilterState = {
    mode: 'all'
  };
  const deferredSaveTimers = new Map();
  let eventListenersBound = false;
  let saveStateTimer = null;
  let isApplyingUndo = false;
  let globalErrorHandlersBound = false;
  const RulesEngine = (typeof window !== 'undefined' && window.SpireRulesEngine) ? window.SpireRulesEngine : {
    getRulesConfig(camp) {
      const defaults = { difficultyDowngrades: true, falloutCheckOnStress: true, clearStressOnFallout: true };
      if (!camp) return defaults;
      const profile = camp.rulesProfile || 'Core';
      if (profile === 'Quickstart') return { difficultyDowngrades: true, falloutCheckOnStress: false, clearStressOnFallout: false };
      if (profile === 'Custom') return Object.assign({}, defaults, camp.customRules || {});
      return defaults;
    },
    totalStressForFallout(pc) {
      const tracks = ['blood', 'mind', 'silver', 'shadow', 'reputation'];
      return tracks.reduce((sum, t) => {
        const filled = (pc.stressFilled && pc.stressFilled[t]) ? pc.stressFilled[t].length : 0;
        let free = 0;
        if (Array.isArray(pc.resistances)) {
          pc.resistances.forEach((r) => {
            if (!r || !r.name) return;
            if (String(r.name).toLowerCase() !== t) return;
            free += Math.max(0, parseInt(r.value, 10) || 0);
          });
        }
        if (t === 'blood' && Array.isArray(pc.inventory)) {
          pc.inventory.forEach((item) => {
            if (!item || item.type !== 'armor') return;
            free += Math.max(0, parseInt(item.resistance, 10) || 0);
          });
        }
        return sum + Math.min(10, Math.max(0, filled - free));
      }, 0);
    },
    falloutSeverityForTotalStress(total) {
      if (total >= 9) return 'Severe';
      if (total >= 5) return 'Moderate';
      return 'Minor';
    },
    stressClearAmountForSeverity(sev) {
      if (sev === 'Severe') return 7;
      if (sev === 'Moderate') return 5;
      return 3;
    }
  };

  /**
   * Generate a unique ID. Combines a prefix with a timestamp and random
   * component to reduce collisions.
   */
  function generateId(prefix = 'id') {
    return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).substr(2, 5)}`;
  }

  /**
   * Create a blank campaign with sensible defaults. Each campaign keeps
   * its own entity and relationship records, log entries and custom
   * relationship types. Positions for the graph are stored per campaign.
   */
  function createCampaign(name = 'New Campaign') {
    const id = generateId('camp');
    return {
      id,
      name,
      schemaVersion: SCHEMA_VERSION,
      nextId: 1,
      gmMode: true,
      gmPin: '',
      allowPlayerEditing: true,
      darkMode: false,
      owner: state.currentUser || null,
      gmUsers: state.currentUser ? [state.currentUser] : [],
      memberUsers: state.currentUser ? [state.currentUser] : [],
      inviteCode: '',
      sourceInviteCode: '',
      relTypes: DEFAULT_REL_TYPES.slice(),
      playerOwnedPcId: null,
      currentSession: 1,
      ministryAttention: 0,   // 0-10 Ministry attention level
      rulesProfile: 'Core',
      entitySort: 'manual',
      entityPinnedOnly: false,
      customRules: {
        difficultyDowngrades: true,
        falloutCheckOnStress: true,
        clearStressOnFallout: true
      },
      falloutGuidance: defaultFalloutGuidance(),
      graphViews: {},
      sectionCollapse: {},    // persistent collapse state per entity+section
      taskViewByPc: {},
      taskSortByPc: {},
      sessionPrepTasksOnly: false,
      entities: {},
      relationships: {},
      positions: {},
      logs: [],
      messages: [],
      gmNotes: [],
      clocks: [],
      lastScenePrompt: null,
      scenarioPacks: [],
      uiTipsDismissed: {},
      relationshipUndo: {},
      relationshipRedo: {},
      undoStack: [],
      redoStack: []
    };
  }

  /**
   * Initialise a new entity of the given type. Populates required
   * attributes based on whether it is a PC, NPC or Organisation.
   */
  function newEntity(campaign, type) {
    const id = generateId(type);
    const base = {
      id,
      type,
      gmOnly: false,
      name: '',
      tags: [],
      notes: '',
      gmNotes: '',
      image: ''
    };
    if (type === 'pc') {
      Object.assign(base, {
        firstName: '',
        lastName: '',
        pronouns: '',
        class: '',
        durance: '',
        classInventorySelections: [], // array of selected kit labels
        coreAbilitiesState: {},
        classBondResponses: [],
        stress: { blood: 0, mind: 0, silver: 0, shadow: 0, reputation: 0 },
        stressSlots:  { blood: 10, mind: 10, silver: 10, shadow: 10, reputation: 10 },
        stressFilled: { blood: [], mind: [], silver: [], shadow: [], reputation: [] },
        advances: [],
        advancePoints: 0,
        fallout: [],
        skills: [],
        domains: [],
        resistances: [],
        inventory: [],
        tasks: [],
        bonds: [],
        refreshed: false
      });
    } else if (type === 'npc') {
      Object.assign(base, {
        role: '',
        affiliation: '',
        pendingApproval: false,
        wants: '',
        fears: '',
        leverage: '',
        threatLevel: 'Minor',
        disposition: 'Neutral',
        bondStressSlots: 10,
        bondStressFilled: [],
        fallout: [],
        inventory: [],
      });
    } else if (type === 'org') {
      Object.assign(base, {
        description: '',
        members: [],
        secrets: [],
        reach: 'Local',
        ministryRelation: 'Unknown'  
      });
    }
    campaign.entities[id] = base;
    return base;
  }

  /**
   * Create a new relationship object between two entities. Relationships
   * default to directed with both parties aware unless specified.
   */
  function newRelationship(campaign, sourceId, targetId, type, options = {}) {
    const id = generateId('rel');
    const rel = {
      id,
      source: sourceId,
      target: targetId,
      type: type || state.relTypes[0],
      directed: options.directed !== undefined ? options.directed : true,
      secret: options.secret || false,
      sourceKnows: options.sourceKnows !== undefined ? options.sourceKnows : true,
      targetKnows: options.targetKnows !== undefined ? options.targetKnows : true,
      notes: options.notes || '',
      // Fallout level for bonds. Defaults to Minor. Can be edited in the
      // relationship inspector. Only relevant when relationships represent bonds.
      falloutLevel: options.falloutLevel || 'Minor'
    };
    campaign.relationships[id] = rel;
    return rel;
  }

  function normalizeStressTrack(track) {
    const t = String(track || '').toLowerCase();
    return ['blood', 'mind', 'silver', 'shadow', 'reputation'].includes(t) ? t : 'mind';
  }

  function applyRelationshipConsequence(rel, opts = {}) {
    const camp = currentCampaign();
    if (!camp || !rel) return false;
    const target = camp.entities[opts.targetId];
    if (!target) return false;
    const source = camp.entities[rel.source];
    const peer = camp.entities[target.id === rel.source ? rel.target : rel.source];
    const relType = rel.type || 'Relationship';
    const severity = opts.severity || rel.falloutLevel || 'Minor';
    const detail = (opts.detail || '').trim();

    if (opts.kind === 'bond-fallout') {
      if (!Array.isArray(target.fallout)) target.fallout = [];
      target.fallout.push({
        id: generateId('fallout'),
        type: opts.track || 'Mind',
        severity,
        name: opts.name || `Bond strain (${relType})`,
        description: detail || `Triggered by ${entityLabel(source)} ${rel.directed ? '->' : '<->'} ${entityLabel(peer || source)}.`,
        resolved: false,
        timestamp: new Date().toISOString()
      });
      appendLog('Applied bond fallout', target.id);
      return true;
    }

    if (opts.kind === 'social-stress') {
      if (target.type !== 'pc') {
        showToast('Social stress can only be applied to PCs.', 'warn');
        return false;
      }
      const track = normalizeStressTrack(opts.track);
      const amount = Math.max(1, Math.min(10, parseInt(opts.amount, 10) || 1));
      const current = (target.stressFilled && target.stressFilled[track]) ? target.stressFilled[track].length : 0;
      const maxSlots = getTrackTotalSlots(target, track);
      setPCStressLevel(target, track, Math.min(maxSlots, current + amount), { triggerFallout: true });
      appendLog('Applied social consequence stress', target.id);
      return true;
    }

    if (opts.kind === 'followup-task') {
      if (target.type !== 'pc') {
        showToast('Follow-up tasks can only be added to PCs.', 'warn');
        return false;
      }
      if (!Array.isArray(target.tasks)) target.tasks = [];
      target.tasks.push({
        id: generateId('task'),
        title: opts.taskTitle || `Follow-up: ${relType} with ${entityLabel(peer || source)}`,
        status: 'To Do',
        priority: 'Normal',
        dueDate: '',
        notes: detail || ''
      });
      appendLog('Created bond follow-up task', target.id);
      return true;
    }

    return false;
  }

  /**
   * Apply effects of a durance to a PC. This removes any previously
   * applied durance modifications (skills, domains, resistances) and
   * applies the bonuses defined in DURANCE_EFFECTS for the given
   * durance. Newly added entries are tagged with a `source`
   * property in the form "durance:<Durance Name>" so that they can
   * be removed later. If `newDurance` is an empty string, any
   * existing durance modifications will be removed and nothing will be
   * added.
   * @param {Object} pc The player character entity
   * @param {String} newDurance The durance key from DURANCE_EFFECTS or ''
   */
  function applyDuranceEffects(pc, newDurance) {
    // Remove existing durance-derived skills/domains/resistances
    pc.skills = pc.skills.filter(s => !s.source || !s.source.startsWith('durance:'));
    pc.domains = pc.domains.filter(d => !d.source || !d.source.startsWith('durance:'));
    pc.resistances = pc.resistances.filter(r => !r.source || !r.source.startsWith('durance:'));
    // Apply new effects if durance selected
    if (!newDurance) return;
    const eff = DURANCE_EFFECTS[newDurance];
    if (!eff) return;
    if (eff.skills) {
      eff.skills.forEach(name => {
        pc.skills.push({ id: generateId('skill'), name, rating: 1, knack: false, source: `durance:${newDurance}` });
      });
    }
    if (eff.domains) {
      eff.domains.forEach(name => {
        pc.domains.push({ id: generateId('domain'), name, source: `durance:${newDurance}` });
      });
    }
    if (eff.resistances) {
      eff.resistances.forEach(({ name, value }) => {
        pc.resistances.push({ id: generateId('res'), name, value, source: `durance:${newDurance}` });
      });
    }
    consolidateResistances(pc);
  }

  /**
   * Apply effects of a class to a PC. Removes previous class-derived
   * skills, domains, resistances and inventory kit items, then applies
   * the new class modifications. The PC's `coreAbilitiesState` and
   * `classInventorySelection` are reset. New entries are tagged with
   * source "class:<Class Name>" or "classInv:<Class Name>".
   * @param {Object} pc The player character entity
   * @param {String} newClass The class key from CLASS_EFFECTS or ''
   */
  function applyClassEffects(pc, newClass) {
    // Remove previous class-derived skills/domains/resistances
    pc.skills = pc.skills.filter(s => !s.source || !s.source.startsWith('class:'));
    pc.domains = pc.domains.filter(d => !d.source || !d.source.startsWith('class:'));
    pc.resistances = pc.resistances.filter(r => !r.source || !r.source.startsWith('class:'));
    // Remove class inventory items
    pc.inventory = pc.inventory.filter(item => !item.source || !item.source.startsWith('classInv:'));
    // Reset core ability usage and class inventory selection
    pc.coreAbilitiesState = {};
    pc.classInventorySelection = '';
    // If no class selected, nothing to add
    if (!newClass) return;
    const eff = CLASS_EFFECTS[newClass];
    if (!eff) return;
    // Apply resistances
    if (eff.resistances) {
      eff.resistances.forEach(({ name, value }) => {
        pc.resistances.push({ id: generateId('res'), name, value, source: `class:${newClass}` });
      });
    }
    // Apply skills
    if (eff.skills) {
      eff.skills.forEach(name => {
        pc.skills.push({ id: generateId('skill'), name, rating: 1, knack: false, source: `class:${newClass}` });
      });
    }
    // Apply domains
    if (eff.domains) {
      eff.domains.forEach(name => {
        pc.domains.push({ id: generateId('domain'), name, source: `class:${newClass}` });
      });
    }
    // Initialise core abilities state (all unused)
    eff.coreAbilities.forEach(ab => {
      pc.coreAbilitiesState[ab] = false;
    });
    // Bond responses: reset length to number of prompts
    pc.classBondResponses = eff.bondPrompts.map(() => '');
    // Note: inventory kits and advances are applied via UI interactions
    consolidateResistances(pc);
  }

  /**
   * Merge resistances that share the same name into a single entry whose
   * value is the sum of all contributions. Sources are joined with ' + '.
   * This is called after applying class and durance effects so overlapping
   * bonuses (e.g. Masked + Information Broker both giving Shadow +2)
   * appear as one row showing Shadow +4 rather than two separate rows.
   */
  function consolidateResistances(pc) {
    const map = {};
    pc.resistances.forEach(r => {
      const key = r.name.toLowerCase();
      if (!map[key]) {
        map[key] = { id: r.id, name: r.name, value: 0, sources: [] };
      }
      map[key].value += (r.value || 0);
      if (r.source) map[key].sources.push(r.source);
    });
    pc.resistances = Object.values(map).map(m => ({
      id: m.id,
      name: m.name,
      value: m.value,
      source: m.sources.join(' + ') || undefined
    }));
  }

  /**
   * Apply a specific inventory kit from the current class to the PC. This
   * removes any previous class inventory items (source starting with
   * "classInv:") and then adds all items from the selected kit. The
   * kitLabel should match one of the labels in CLASS_EFFECTS[pc.class].
   * @param {Object} pc The player character entity
   * @param {String} kitLabel The label of the inventory option selected
   */
  function applyClassInventoryKit(pc, kitLabel) {
    // Remove all previous class inventory items
    pc.inventory = pc.inventory.filter(it => !it.source || !it.source.startsWith('classInv:'));
    pc.classInventorySelection = kitLabel;
    const classEff = CLASS_EFFECTS[pc.class];
    if (!classEff || !classEff.inventoryOptions) return;
    const option = classEff.inventoryOptions.find(opt => opt.label === kitLabel);
    if (!option) return;
    option.items.forEach(item => {
      const newItem = {
        id: generateId('item'),
        item: item.item,
        quantity: item.quantity || 1,
        type: item.type || 'other',
        tags: item.tags ? item.tags.slice() : [],
        notes: '',
        stress: item.stress,
        resistance: item.resistance,
        source: `classInv:${pc.class}`
      };
      pc.inventory.push(newItem);
    });
  }

  /**
   * Load campaigns from localStorage. If none exist, initialise a new
   * default campaign. Returns true if campaigns loaded successfully.
   */
  function userScopedKey(base) {
    return state.currentUser ? `${base}:${state.currentUser}` : base;
  }

  function syncConfigStorageKey() {
    return userScopedKey('spire-sync-config');
  }

  function telemetryConfigStorageKey() {
    return userScopedKey('spire-telemetry-config');
  }

  function crashLogStorageKey() {
    return userScopedKey('spire-crash-log');
  }

  function defaultSyncConfig() {
    const bootUrl = (window.SPIRE_SUPABASE_URL || '').trim();
    const bootKey = (window.SPIRE_SUPABASE_ANON_KEY || '').trim();
    return {
      transport: (bootUrl && bootKey) ? 'supabase' : 'local',
      supabaseUrl: bootUrl,
      supabaseAnonKey: bootKey
    };
  }

  function defaultTelemetryConfig() {
    return {
      enabled: false,
      endpoint: '',
      maxLocal: 50
    };
  }

  function loadSyncConfig() {
    try {
      const raw = localStorage.getItem(syncConfigStorageKey());
      if (!raw) return defaultSyncConfig();
      const parsed = JSON.parse(raw) || {};
      const cfg = defaultSyncConfig();
      if (parsed.transport === 'supabase') cfg.transport = 'supabase';
      if (typeof parsed.supabaseUrl === 'string') cfg.supabaseUrl = parsed.supabaseUrl.trim();
      if (typeof parsed.supabaseAnonKey === 'string') cfg.supabaseAnonKey = parsed.supabaseAnonKey.trim();
      return cfg;
    } catch (_) {
      return defaultSyncConfig();
    }
  }

  function loadTelemetryConfig() {
    try {
      const raw = localStorage.getItem(telemetryConfigStorageKey());
      if (!raw) return defaultTelemetryConfig();
      const parsed = JSON.parse(raw) || {};
      return {
        enabled: !!parsed.enabled,
        endpoint: typeof parsed.endpoint === 'string' ? parsed.endpoint.trim() : '',
        maxLocal: Math.max(10, Math.min(200, parseInt(parsed.maxLocal, 10) || 50))
      };
    } catch (_) {
      return defaultTelemetryConfig();
    }
  }

  function saveTelemetryConfig(config = {}) {
    const safe = defaultTelemetryConfig();
    safe.enabled = !!config.enabled;
    safe.endpoint = typeof config.endpoint === 'string' ? config.endpoint.trim() : '';
    safe.maxLocal = Math.max(10, Math.min(200, parseInt(config.maxLocal, 10) || 50));
    state.telemetryConfig = safe;
    try {
      localStorage.setItem(telemetryConfigStorageKey(), JSON.stringify(safe));
    } catch (_) {}
    return safe;
  }

  function loadCrashLog() {
    try {
      const raw = localStorage.getItem(crashLogStorageKey());
      const parsed = raw ? JSON.parse(raw) : [];
      return Array.isArray(parsed) ? parsed : [];
    } catch (_) {
      return [];
    }
  }

  function saveCrashLog(logs) {
    try {
      localStorage.setItem(crashLogStorageKey(), JSON.stringify(Array.isArray(logs) ? logs : []));
    } catch (_) {}
  }

  function enqueueCrashReport(report = {}) {
    const cfg = state.telemetryConfig || loadTelemetryConfig();
    const entry = Object.assign({
      id: generateId('crash'),
      time: new Date().toISOString(),
      user: state.currentUser || '',
      campaignId: state.currentCampaignId || '',
      gmMode: !!state.gmMode,
      userAgent: navigator.userAgent || ''
    }, report || {});
    const logs = loadCrashLog();
    logs.push(entry);
    while (logs.length > (cfg.maxLocal || 50)) logs.shift();
    saveCrashLog(logs);

    if (!cfg.enabled || !cfg.endpoint) return;
    try {
      const body = JSON.stringify(entry);
      if (navigator.sendBeacon) {
        const blob = new Blob([body], { type: 'application/json' });
        navigator.sendBeacon(cfg.endpoint, blob);
      } else {
        fetch(cfg.endpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body
        }).catch(() => {});
      }
    } catch (_) {}
  }

  function bindGlobalErrorHandlers() {
    if (globalErrorHandlersBound) return;
    globalErrorHandlersBound = true;
    window.addEventListener('error', (event) => {
      try {
        enqueueCrashReport({
          kind: 'error',
          message: event.message || 'Unknown error',
          source: event.filename || '',
          line: Number(event.lineno) || 0,
          column: Number(event.colno) || 0,
          stack: event.error && event.error.stack ? String(event.error.stack) : ''
        });
      } catch (_) {}
    });
    window.addEventListener('unhandledrejection', (event) => {
      try {
        const reason = event.reason;
        enqueueCrashReport({
          kind: 'unhandledrejection',
          message: typeof reason === 'string' ? reason : (reason && reason.message) ? reason.message : 'Unhandled rejection',
          stack: reason && reason.stack ? String(reason.stack) : ''
        });
      } catch (_) {}
    });
  }

  function saveSyncConfig(config = {}) {
    const safe = defaultSyncConfig();
    if (config.transport === 'supabase') safe.transport = 'supabase';
    if (typeof config.supabaseUrl === 'string') safe.supabaseUrl = config.supabaseUrl.trim();
    if (typeof config.supabaseAnonKey === 'string') safe.supabaseAnonKey = config.supabaseAnonKey.trim();
    state.syncConfig = safe;
    try {
      localStorage.setItem(syncConfigStorageKey(), JSON.stringify(safe));
    } catch (_) {
      // ignore storage errors
    }
    return safe;
  }

  function loadUsers() {
    try {
      const raw = localStorage.getItem('spire-users');
      state.users = raw ? (JSON.parse(raw) || {}) : {};
      const current = localStorage.getItem('spire-current-user');
      const allowCloudUser = !!(current && canUseOnlineApi());
      state.currentUser = (current && (state.users[current] || allowCloudUser)) ? current : null;
      if (!state.currentUser) localStorage.removeItem('spire-current-user');
      return true;
    } catch (e) {
      console.error('Failed to load users', e);
      state.users = {};
      state.currentUser = null;
      return false;
    }
  }

  function saveUsers() {
    try {
      localStorage.setItem('spire-users', JSON.stringify(state.users));
      if (state.currentUser) localStorage.setItem('spire-current-user', state.currentUser);
      else localStorage.removeItem('spire-current-user');
    } catch (e) {
      console.warn('Failed to save users', e);
    }
  }

  function loadSharedInvites() {
    try {
      const raw = localStorage.getItem('spire-shared-invites');
      return raw ? (JSON.parse(raw) || {}) : {};
    } catch (_) {
      return {};
    }
  }

  function saveSharedInvites(data) {
    try {
      localStorage.setItem('spire-shared-invites', JSON.stringify(data || {}));
    } catch (e) {
      console.warn('Failed to save shared invites', e);
    }
  }

  function generateInviteCode() {
    const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    let out = '';
    for (let i = 0; i < 8; i += 1) out += alphabet[Math.floor(Math.random() * alphabet.length)];
    return out;
  }

  function publishCampaignInvite(camp) {
    if (!camp || !camp.inviteCode || !camp.owner) return;
    const code = String(camp.inviteCode).trim().toUpperCase();
    if (!code) return;
    if (!Array.isArray(camp.memberUsers)) camp.memberUsers = camp.owner ? [camp.owner] : [];
    if (camp.owner && !camp.memberUsers.includes(camp.owner)) camp.memberUsers.unshift(camp.owner);
    if (!Array.isArray(camp.gmUsers)) camp.gmUsers = camp.owner ? [camp.owner] : [];
    if (camp.owner && !camp.gmUsers.includes(camp.owner)) camp.gmUsers.unshift(camp.owner);

    const shared = loadSharedInvites();
    const existing = shared[code];
    const existingMembers = Array.isArray(existing?.data?.memberUsers) ? existing.data.memberUsers : [];
    camp.memberUsers = Array.from(new Set([...(camp.memberUsers || []), ...existingMembers]));

    const payload = JSON.parse(JSON.stringify(camp));
    payload.undoStack = [];
    payload.redoStack = [];
    shared[code] = {
      code,
      owner: camp.owner,
      name: camp.name,
      updatedAt: new Date().toISOString(),
      data: payload
    };
    saveSharedInvites(shared);
  }

  function revokeCampaignInvite(camp) {
    if (!camp || !camp.inviteCode) return;
    const code = String(camp.inviteCode).trim().toUpperCase();
    const shared = loadSharedInvites();
    delete shared[code];
    saveSharedInvites(shared);
  }

  function syncOwnedSharedInvites() {
    if (!state.currentUser) return;
    const shared = loadSharedInvites();
    // Remove stale records owned by current user when campaign/code is gone.
    Object.keys(shared).forEach((code) => {
      const rec = shared[code];
      if (!rec || rec.owner !== state.currentUser) return;
      const stillExists = Object.values(state.campaigns || {}).some((camp) =>
        camp.owner === state.currentUser && String(camp.inviteCode || '').toUpperCase() === code
      );
      if (!stillExists) delete shared[code];
    });
    saveSharedInvites(shared);
    Object.values(state.campaigns || {}).forEach((camp) => {
      if (camp.owner === state.currentUser && camp.inviteCode) publishCampaignInvite(camp);
    });
  }

  async function hashPassword(password) {
    const text = String(password || '');
    if (window.crypto && window.crypto.subtle && window.TextEncoder) {
      const data = new TextEncoder().encode(text);
      const digest = await window.crypto.subtle.digest('SHA-256', data);
      return Array.from(new Uint8Array(digest)).map(b => b.toString(16).padStart(2, '0')).join('');
    }
    return btoa(unescape(encodeURIComponent(text)));
  }

  function loadCampaigns() {
    try {
      if (!state.currentUser) {
        state.campaigns = {};
        state.currentCampaignId = null;
        state.lastSeenRevision = '';
        return true;
      }
      state.lastSeenRevision = localStorage.getItem(userScopedKey('spire-campaigns-rev')) || '';
      const raw = localStorage.getItem(userScopedKey('spire-campaigns'));
      if (raw) {
        const parsed = JSON.parse(raw);
        // Migrate if needed (schemaVersion future use)
        state.campaigns = parsed.campaigns || {};
        state.currentCampaignId = parsed.currentCampaignId;
        // Sanity check: ensure at least one campaign exists
        if (!state.currentCampaignId || !state.campaigns[state.currentCampaignId]) {
          const camp = createCampaign('Default');
          state.campaigns[camp.id] = camp;
          state.currentCampaignId = camp.id;
        }
      } else {
        // Create a default campaign if none exist
        const camp = createCampaign('Default');
        state.campaigns[camp.id] = camp;
        state.currentCampaignId = camp.id;
      }
      return true;
    } catch (e) {
      console.error('Failed to load campaigns', e);
      return false;
    }
  }

  /**
   * Save campaigns and current campaign id back into localStorage.
   */
  function conflictWarningText() {
    const edits = Number(state.localEditsSinceConflict || 0);
    if (!state.syncConflictActive) return 'Updated in another tab';
    if (edits <= 0) return 'Updated in another tab';
    return `Updated in another tab (${edits} local change${edits === 1 ? '' : 's'} pending)`;
  }

  function pendingOpsStorageKey() {
    return userScopedKey('spire-pending-ops');
  }

  function loadPendingOps() {
    try {
      const raw = localStorage.getItem(pendingOpsStorageKey());
      const parsed = raw ? JSON.parse(raw) : [];
      return Array.isArray(parsed) ? parsed : [];
    } catch (_) {
      return [];
    }
  }

  function savePendingOps(ops) {
    try {
      localStorage.setItem(pendingOpsStorageKey(), JSON.stringify(Array.isArray(ops) ? ops : []));
    } catch (_) {
      // ignore storage failures
    }
    updateSyncQueueButton();
  }

  function removePendingOpsById(ids = []) {
    const remove = new Set((Array.isArray(ids) ? ids : []).filter(Boolean));
    if (!remove.size) return;
    const next = loadPendingOps().filter((op) => !remove.has(op.id));
    savePendingOps(next);
  }

  function enqueuePendingSaveOp() {
    if (!state.currentUser) return;
    const data = {
      campaigns: state.campaigns,
      currentCampaignId: state.currentCampaignId
    };
    const ops = loadPendingOps();
    ops.push({
      id: generateId('op'),
      time: new Date().toISOString(),
      kind: 'saveCampaigns',
      baseRevision: state.lastSeenRevision || localStorage.getItem(userScopedKey('spire-campaigns-rev')) || '',
      payload: data
    });
    // Keep queue bounded.
    while (ops.length > 50) ops.shift();
    savePendingOps(ops);
  }

  function flushPendingSaveOps() {
    if (!navigator.onLine || !state.currentUser) return false;
    const ops = loadPendingOps();
    if (!ops.length) return false;
    const latest = ops[ops.length - 1];
    if (!latest || latest.kind !== 'saveCampaigns' || !latest.payload) return false;
    try {
      const currentRev = localStorage.getItem(userScopedKey('spire-campaigns-rev')) || '';
      if (latest.baseRevision && currentRev && latest.baseRevision !== currentRev) {
        const msg = `Sync conflict: queued changes based on older revision (${ops.length} queued)`;
        setSyncConflictWarning(true, msg);
        setSaveState('error', msg);
        return false;
      }
      const serialized = JSON.stringify(latest.payload);
      localStorage.setItem(userScopedKey('spire-campaigns'), serialized);
      const revisionToken = `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
      localStorage.setItem(userScopedKey('spire-campaigns-rev'), revisionToken);
      state.lastSeenRevision = revisionToken;
      localStorage.setItem(userScopedKey('spire-campaigns-backup'), serialized);
      localStorage.setItem(userScopedKey('spire-campaigns-backup-ts'), new Date().toISOString());
      savePendingOps([]);
      broadcastRealtimeUpdate({
        type: 'campaign_saved',
        revision: revisionToken,
        campaignId: state.currentCampaignId,
        actor: currentActorLabel(),
        actorRole: currentActorRole()
      });
      setSaveState('saved', 'Synced queued changes');
      return true;
    } catch (_) {
      return false;
    }
  }

  function closeRealtimeChannel() {
    try {
      if (state.realtimeChannel) state.realtimeChannel.close();
    } catch (_) {}
    try {
      if (state.supabaseRealtimeChannel) {
        if (typeof state.supabaseRealtimeChannel.unsubscribe === 'function') {
          state.supabaseRealtimeChannel.unsubscribe();
        } else if (typeof state.supabaseRealtimeChannel.close === 'function') {
          state.supabaseRealtimeChannel.close();
        }
      }
    } catch (_) {}
    state.realtimeChannel = null;
    state.supabaseRealtimeChannel = null;
    state.realtimeTransport = 'local';
  }

  function handleIncomingRealtimeMessage(msg) {
    if (!msg || msg.clientId === state.clientId) return;
    if (msg.type !== 'campaign_saved') return;
    if (msg.revision && msg.revision === state.lastSeenRevision) return;
    state.lastSeenRevision = msg.revision || state.lastSeenRevision;
    const actor = msg.actor ? ` by ${msg.actor}` : '';
    const conflictMsg = `Updated elsewhere${actor}`;
    setSyncConflictWarning(true, conflictMsg);
    setSaveState('error', conflictMsg);
  }

  function canUseSupabaseRealtime() {
    const cfg = state.syncConfig || loadSyncConfig();
    if (!cfg || cfg.transport !== 'supabase') return false;
    return !!(cfg.supabaseUrl && cfg.supabaseAnonKey && window.SpireOnlineClient && typeof window.SpireOnlineClient.init === 'function');
  }

  function canUseOnlineApi() {
    const cfg = state.syncConfig || loadSyncConfig();
    return !!(cfg && cfg.transport === 'supabase' && cfg.supabaseUrl && cfg.supabaseAnonKey
      && window.SpireOnlineClient && typeof window.SpireOnlineClient.init === 'function');
  }

  function initOnlineClient() {
    if (!canUseOnlineApi()) return false;
    try {
      const cfg = state.syncConfig || loadSyncConfig();
      window.SpireOnlineClient.init({
        url: cfg.supabaseUrl,
        anonKey: cfg.supabaseAnonKey
      });
      return true;
    } catch (e) {
      console.warn('Online client init failed', e);
      return false;
    }
  }

  function usernameToAuthEmail(username) {
    const raw = String(username || '').trim().toLowerCase();
    if (!raw) return '';
    if (raw.includes('@')) return raw;
    const safe = raw.replace(/[^a-z0-9._-]/g, '');
    return `${safe || 'user'}@spire.local`;
  }

  function usernameFromOnlineUser(user, fallback = '') {
    const metaName = user && user.user_metadata && typeof user.user_metadata.username === 'string'
      ? user.user_metadata.username.trim()
      : '';
    if (metaName) return metaName;
    const email = user && typeof user.email === 'string' ? user.email.trim() : '';
    if (email.includes('@')) return email.split('@')[0];
    return fallback || '';
  }

  async function syncCampaignsFromCloud() {
    if (!state.currentUser || !canUseOnlineApi() || !initOnlineClient()) return false;
    const rows = await window.SpireOnlineClient.listMyCampaigns();
    const nextCampaigns = {};
    (rows || []).forEach((row) => {
      const remote = row && row.campaign ? row.campaign : null;
      if (!remote || !remote.id) return;
      const remoteData = (remote.data && typeof remote.data === 'object') ? JSON.parse(JSON.stringify(remote.data)) : {};
      const localId = (typeof remoteData.id === 'string' && remoteData.id.trim()) ? remoteData.id.trim() : String(remote.id);
      if (!remoteData.entities || typeof remoteData.entities !== 'object') remoteData.entities = {};
      if (!remoteData.relationships || typeof remoteData.relationships !== 'object') remoteData.relationships = {};
      if (!Array.isArray(remoteData.logs)) remoteData.logs = [];
      if (!Array.isArray(remoteData.relTypes)) remoteData.relTypes = DEFAULT_REL_TYPES.slice();
      if (!Array.isArray(remoteData.memberUsers)) remoteData.memberUsers = [];
      if (!remoteData.memberUsers.includes(state.currentUser)) remoteData.memberUsers.push(state.currentUser);
      if (!Array.isArray(remoteData.gmUsers)) remoteData.gmUsers = [];
      if (row && row.role === 'gm' && !remoteData.gmUsers.includes(state.currentUser)) remoteData.gmUsers.push(state.currentUser);
      remoteData.id = localId;
      remoteData.cloudCampaignId = String(remote.id);
      remoteData.name = remoteData.name || remote.name || 'Campaign';
      if (!remoteData.owner) remoteData.owner = state.currentUser;
      nextCampaigns[localId] = remoteData;
    });
    if (!Object.keys(nextCampaigns).length) return false;
    state.campaigns = nextCampaigns;
    if (!state.currentCampaignId || !state.campaigns[state.currentCampaignId]) {
      state.currentCampaignId = Object.keys(state.campaigns)[0];
    }
    saveCampaigns({ force: true });
    return true;
  }

  async function persistCurrentCampaignToCloud() {
    if (!state.currentUser || !canUseOnlineApi() || !initOnlineClient()) return false;
    const camp = currentCampaign();
    if (!camp) return false;
    let cloudId = camp.cloudCampaignId || '';
    if (!cloudId) {
      const createdId = await window.SpireOnlineClient.createCampaign(camp.name || 'New Campaign');
      cloudId = String(createdId || '').trim();
      if (!cloudId) throw new Error('Failed to create cloud campaign.');
      camp.cloudCampaignId = cloudId;
    }
    await window.SpireOnlineClient.saveCampaignData(cloudId, camp);
    return true;
  }

  function initSupabaseRealtimeChannel() {
    if (!state.currentUser || !canUseSupabaseRealtime()) return false;
    try {
      const cfg = state.syncConfig || loadSyncConfig();
      const client = window.SpireOnlineClient.init({
        url: cfg.supabaseUrl,
        anonKey: cfg.supabaseAnonKey
      });
      if (!client || typeof client.channel !== 'function') return false;
      const channel = client.channel(`spire-campaigns-${state.currentUser}`);
      channel
        .on('broadcast', { event: 'campaign_saved' }, (evt) => {
          const payload = evt && evt.payload ? evt.payload : null;
          handleIncomingRealtimeMessage(payload);
        })
        .subscribe();
      state.supabaseRealtimeChannel = channel;
      state.realtimeTransport = 'supabase';
      return true;
    } catch (e) {
      console.warn('Supabase realtime disabled; falling back to local transport.', e);
      state.supabaseRealtimeChannel = null;
      state.realtimeTransport = 'local';
      return false;
    }
  }

  function initRealtimeChannel() {
    closeRealtimeChannel();
    if (!state.syncConfig) state.syncConfig = loadSyncConfig();
    if (!state.currentUser) return;
    if (initSupabaseRealtimeChannel()) return;
    if (typeof BroadcastChannel === 'undefined') return;
    try {
      const channel = new BroadcastChannel(`spire-realtime-${state.currentUser}`);
      channel.addEventListener('message', (evt) => {
        const msg = evt && evt.data ? evt.data : null;
        handleIncomingRealtimeMessage(msg);
      });
      state.realtimeChannel = channel;
      state.realtimeTransport = 'local';
    } catch (_) {
      state.realtimeChannel = null;
      state.realtimeTransport = 'local';
    }
  }

  function broadcastRealtimeUpdate(payload = {}) {
    const msg = Object.assign({}, payload, { clientId: state.clientId, time: new Date().toISOString() });
    if (state.realtimeTransport === 'supabase' && state.supabaseRealtimeChannel && typeof state.supabaseRealtimeChannel.send === 'function') {
      try {
        state.supabaseRealtimeChannel.send({
          type: 'broadcast',
          event: payload.type || 'campaign_saved',
          payload: msg
        });
      } catch (_) {
        // ignore and fall back to local channel below
      }
    }
    const ch = state.realtimeChannel;
    if (!ch) return;
    try {
      ch.postMessage(msg);
    } catch (_) {
      // ignore
    }
  }

  function saveCampaigns(options = {}) {
    const force = !!options.force;
    if (state.syncConflictActive && !force) {
      state.localEditsSinceConflict = (state.localEditsSinceConflict || 0) + 1;
      const conflictMsg = conflictWarningText();
      setSyncConflictWarning(true, conflictMsg);
      setSaveState('error', conflictMsg);
      return false;
    }
    setSaveState('saving');
    try {
      if (!navigator.onLine) {
        enqueuePendingSaveOp();
        setSaveState('error', 'Offline (queued)');
        return false;
      }
      if (!state.currentUser) {
        setSaveState('saved');
        return true;
      }
      const data = {
        campaigns: state.campaigns,
        currentCampaignId: state.currentCampaignId
      };
      const serialized = JSON.stringify(data);
      localStorage.setItem(userScopedKey('spire-campaigns'), serialized);
      const revisionToken = `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
      localStorage.setItem(userScopedKey('spire-campaigns-rev'), revisionToken);
      state.lastSeenRevision = revisionToken;
      // Keep a rolling backup for quick manual recovery.
      localStorage.setItem(userScopedKey('spire-campaigns-backup'), serialized);
      localStorage.setItem(userScopedKey('spire-campaigns-backup-ts'), new Date().toISOString());
      syncOwnedSharedInvites();
      broadcastRealtimeUpdate({
        type: 'campaign_saved',
        revision: revisionToken,
        campaignId: state.currentCampaignId,
        actor: currentActorLabel(),
        actorRole: currentActorRole()
      });
      if (canUseOnlineApi()) {
        setTimeout(() => {
          persistCurrentCampaignToCloud().catch((err) => {
            console.warn('Cloud save failed', err);
            setSaveState('error', 'Cloud save failed');
          });
        }, 0);
      }
      setSaveState('saved');
      setSyncConflictWarning(false);
      return true;
    } catch (e) {
      console.warn('Save failed', e);
      setSaveState('error');
      return false;
    }
  }

  /**
   * Get the currently active campaign.
   */
  function currentCampaign() {
    return state.campaigns[state.currentCampaignId];
  }

  function getRulesConfig(camp = currentCampaign()) {
    return RulesEngine.getRulesConfig(camp);
  }

  /**
   * Append a log entry to the campaign's log. In player mode, sensitive
   * actions are omitted from the log view (but still recorded for GM).
   */
  // High-frequency stress actions filtered from log to reduce noise
  const LOG_STRESS_PREFIXES = ['Set blood stress', 'Set mind stress', 'Set silver stress',
    'Set shadow stress', 'Set reputation stress', 'Set bond stress'];

  function formatRelationshipLabel(camp, rel) {
    if (!rel) return '';
    const source = camp.entities[rel.source];
    const target = camp.entities[rel.target];
    const sourceLabel = source ? entityLabel(source) : 'Unknown';
    const targetLabel = target ? entityLabel(target) : 'Unknown';
    const relType = rel.type || 'Relationship';
    return `${sourceLabel} ${relType} ${targetLabel}`;
  }

  function describeLogTarget(camp, targetId) {
    if (!targetId) return '';
    const ent = camp.entities[targetId];
    if (ent) return entityLabel(ent);
    const rel = camp.relationships[targetId];
    if (rel) return formatRelationshipLabel(camp, rel);
    return '';
  }

  function resolveLogTargetLabel(camp, entry) {
    if (!entry || !entry.target) return '';
    const live = describeLogTarget(camp, entry.target);
    return live || entry.targetLabel || '';
  }

  function currentActorLabel() {
    return state.currentUser || (state.gmMode ? 'GM' : 'Player');
  }

  function currentActorRole() {
    return state.gmMode ? 'gm' : 'player';
  }

  function appendLog(action, entityOrRelId, type = 'action') {
    for (const prefix of LOG_STRESS_PREFIXES) {
      if (action.startsWith(prefix)) return;
    }
    const camp = currentCampaign();
    const targetLabel = describeLogTarget(camp, entityOrRelId);
    camp.logs.push({
      time: new Date().toISOString(),
      action,
      target: entityOrRelId,
      targetLabel,
      actor: currentActorLabel(),
      actorRole: currentActorRole(),
      type,
      session: camp.currentSession || 1
    });
  }

  function appendSessionLog(label) {
    const camp = currentCampaign();
    camp.logs.push({
      time: new Date().toISOString(),
      action: label,
      target: '',
      actor: currentActorLabel(),
      actorRole: currentActorRole(),
      type: 'session',
      session: camp.currentSession || 1
    });
  }

  function getRecentEntityActions(entityId, limit = 12) {
    const camp = currentCampaign();
    if (!entityId || !camp || !Array.isArray(camp.logs)) return [];
    return camp.logs
      .filter((entry) => entry && entry.type !== 'session' && entry.target === entityId)
      .slice()
      .reverse()
      .slice(0, limit);
  }

  function resolveUndoTargetLabel(targetId, camp = currentCampaign()) {
    if (!targetId || !camp) return '';
    const ent = camp.entities ? camp.entities[targetId] : null;
    if (ent) return entityLabel(ent);
    const rel = camp.relationships ? camp.relationships[targetId] : null;
    if (rel) {
      const src = camp.entities[rel.source];
      const trg = camp.entities[rel.target];
      return `${entityLabel(src)} -> ${entityLabel(trg)}`;
    }
    return '';
  }

  function appendGettingStartedHelper(host, key, helpText) {
    const camp = currentCampaign();
    if (!host || !camp) return;
    if (!camp.uiTipsDismissed || typeof camp.uiTipsDismissed !== 'object') camp.uiTipsDismissed = {};
    const tipKey = `help:${key}`;
    if (camp.uiTipsDismissed[tipKey]) return;
    const wrap = document.createElement('details');
    wrap.className = 'starter-help';
    const summary = document.createElement('summary');
    summary.textContent = 'Getting started';
    const body = document.createElement('p');
    body.textContent = helpText;
    const dismiss = document.createElement('button');
    dismiss.className = 'toolbar-btn';
    dismiss.textContent = 'Hide this tip';
    dismiss.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      camp.uiTipsDismissed[tipKey] = true;
      saveCampaigns();
      wrap.remove();
    });
    wrap.appendChild(summary);
    wrap.appendChild(body);
    wrap.appendChild(dismiss);
    host.appendChild(wrap);
  }

  function updateUndoButtonState() {
    const btn = document.getElementById('undo-btn');
    const camp = currentCampaign();
    const hasUndo = !!(camp && Array.isArray(camp.undoStack) && camp.undoStack.length);
    if (btn) {
      btn.disabled = !hasUndo;
      btn.title = hasUndo
        ? `Undo: ${camp.undoStack[camp.undoStack.length - 1].label || 'last action'}`
        : 'Nothing to undo';
    }
    updateRedoButtonState();
  }

  function updateRedoButtonState() {
    const btn = document.getElementById('redo-btn');
    if (!btn) return;
    const camp = currentCampaign();
    const hasRedo = !!(camp && Array.isArray(camp.redoStack) && camp.redoStack.length);
    btn.disabled = !hasRedo;
    btn.title = hasRedo
      ? `Redo: ${camp.redoStack[camp.redoStack.length - 1].label || 'last action'}`
      : 'Nothing to redo';
  }

  function captureUndoSnapshot(label = 'Change', targetId = '') {
    if (isApplyingUndo) return;
    const camp = currentCampaign();
    if (!camp) return;
    if (!Array.isArray(camp.undoStack)) camp.undoStack = [];
    if (!Array.isArray(camp.redoStack)) camp.redoStack = [];
    const snapshot = JSON.parse(JSON.stringify(camp));
    snapshot.undoStack = [];
    snapshot.redoStack = [];
    const resolvedTarget = targetId || state.selectedEntityId || state.selectedRelId || '';
    const targetLabel = resolveUndoTargetLabel(resolvedTarget, camp);
    camp.undoStack.push({
      id: generateId('undo'),
      time: new Date().toISOString(),
      label,
      target: resolvedTarget,
      targetLabel,
      snapshot
    });
    // New action invalidates redo history.
    camp.redoStack = [];
    if (camp.undoStack.length > 20) camp.undoStack.shift();
    updateUndoButtonState();
  }

  function ensureRelationshipUndoStore(camp = currentCampaign()) {
    if (!camp.relationshipUndo || typeof camp.relationshipUndo !== 'object') camp.relationshipUndo = {};
    if (!camp.relationshipRedo || typeof camp.relationshipRedo !== 'object') camp.relationshipRedo = {};
  }

  function pushRelationshipUndo(relId, label = 'Edit relationship') {
    const camp = currentCampaign();
    const rel = camp && camp.relationships ? camp.relationships[relId] : null;
    if (!camp || !rel) return;
    ensureRelationshipUndoStore(camp);
    if (!Array.isArray(camp.relationshipUndo[relId])) camp.relationshipUndo[relId] = [];
    camp.relationshipUndo[relId].push({
      id: generateId('rundo'),
      time: new Date().toISOString(),
      label,
      data: JSON.parse(JSON.stringify(rel))
    });
    // New relationship edit invalidates redo for this relationship.
    camp.relationshipRedo[relId] = [];
    if (camp.relationshipUndo[relId].length > 30) camp.relationshipUndo[relId].shift();
  }

  function undoRelationshipEdit(relId) {
    const camp = currentCampaign();
    const rel = camp && camp.relationships ? camp.relationships[relId] : null;
    if (!camp || !rel) return;
    ensureRelationshipUndoStore(camp);
    const stack = camp.relationshipUndo[relId];
    if (!Array.isArray(stack) || !stack.length) {
      showToast('No relationship edits to undo.', 'warn');
      return;
    }
    const entry = stack.pop();
    if (!Array.isArray(camp.relationshipRedo[relId])) camp.relationshipRedo[relId] = [];
    camp.relationshipRedo[relId].push({
      id: generateId('rredo'),
      time: new Date().toISOString(),
      label: entry.label || 'relationship edit',
      data: JSON.parse(JSON.stringify(rel))
    });
    if (camp.relationshipRedo[relId].length > 30) camp.relationshipRedo[relId].shift();
    camp.relationships[relId] = JSON.parse(JSON.stringify(entry.data || rel));
    appendLog('Undid relationship edit', relId);
    saveAndRefresh();
    selectRelationship(relId);
  }

  function redoRelationshipEdit(relId) {
    const camp = currentCampaign();
    const rel = camp && camp.relationships ? camp.relationships[relId] : null;
    if (!camp || !rel) return;
    ensureRelationshipUndoStore(camp);
    const redo = camp.relationshipRedo[relId];
    if (!Array.isArray(redo) || !redo.length) {
      showToast('No relationship edits to redo.', 'warn');
      return;
    }
    if (!Array.isArray(camp.relationshipUndo[relId])) camp.relationshipUndo[relId] = [];
    camp.relationshipUndo[relId].push({
      id: generateId('rundo'),
      time: new Date().toISOString(),
      label: redo[redo.length - 1].label || 'relationship edit',
      data: JSON.parse(JSON.stringify(rel))
    });
    if (camp.relationshipUndo[relId].length > 30) camp.relationshipUndo[relId].shift();
    const entry = redo.pop();
    camp.relationships[relId] = JSON.parse(JSON.stringify(entry.data || rel));
    appendLog('Redid relationship edit', relId);
    saveAndRefresh();
    selectRelationship(relId);
  }

  function ensureSectionUndoStore(ent) {
    if (!ent || typeof ent !== 'object') return;
    if (!ent.sectionUndo || typeof ent.sectionUndo !== 'object') ent.sectionUndo = {};
    if (!ent.sectionRedo || typeof ent.sectionRedo !== 'object') ent.sectionRedo = {};
    ['tasks', 'inventory', 'bonds'].forEach((k) => {
      if (!Array.isArray(ent.sectionUndo[k])) ent.sectionUndo[k] = [];
      if (!Array.isArray(ent.sectionRedo[k])) ent.sectionRedo[k] = [];
    });
  }

  function pushSectionUndo(ent, sectionKey, label = 'Edit section') {
    if (!ent || !sectionKey) return;
    ensureSectionUndoStore(ent);
    const sectionData = Array.isArray(ent[sectionKey]) ? ent[sectionKey] : [];
    ent.sectionUndo[sectionKey].push({
      id: generateId('sundo'),
      time: new Date().toISOString(),
      label,
      data: JSON.parse(JSON.stringify(sectionData))
    });
    // New edit invalidates redo for this section.
    ent.sectionRedo[sectionKey] = [];
    if (ent.sectionUndo[sectionKey].length > 20) ent.sectionUndo[sectionKey].shift();
  }

  function undoSection(ent, sectionKey, emptyMsg = 'Nothing to undo.') {
    if (!ent || !sectionKey) return;
    ensureSectionUndoStore(ent);
    const stack = ent.sectionUndo[sectionKey];
    if (!stack.length) {
      showToast(emptyMsg, 'warn');
      return;
    }
    const currentData = Array.isArray(ent[sectionKey]) ? ent[sectionKey] : [];
    ent.sectionRedo[sectionKey].push({
      id: generateId('sredo'),
      time: new Date().toISOString(),
      label: stack[stack.length - 1]?.label || 'section edit',
      data: JSON.parse(JSON.stringify(currentData))
    });
    if (ent.sectionRedo[sectionKey].length > 20) ent.sectionRedo[sectionKey].shift();
    const entry = stack.pop();
    ent[sectionKey] = JSON.parse(JSON.stringify(entry.data || []));
    appendLog(`Undid ${sectionKey} change`, ent.id);
    saveAndRefresh();
  }

  function sectionUndoTitle(ent, sectionKey, emptyMsg = 'Nothing to undo.') {
    ensureSectionUndoStore(ent);
    const stack = ent.sectionUndo[sectionKey];
    if (!stack || !stack.length) return emptyMsg;
    const last = stack[stack.length - 1];
    const when = last.time ? new Date(last.time).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '';
    return `Undo: ${last.label || 'last change'}${when ? ` (${when})` : ''}`;
  }

  function redoSection(ent, sectionKey, emptyMsg = 'Nothing to redo.') {
    if (!ent || !sectionKey) return;
    ensureSectionUndoStore(ent);
    const stack = ent.sectionRedo[sectionKey];
    if (!stack.length) {
      showToast(emptyMsg, 'warn');
      return;
    }
    const currentData = Array.isArray(ent[sectionKey]) ? ent[sectionKey] : [];
    ent.sectionUndo[sectionKey].push({
      id: generateId('sundo'),
      time: new Date().toISOString(),
      label: stack[stack.length - 1]?.label || 'section edit',
      data: JSON.parse(JSON.stringify(currentData))
    });
    if (ent.sectionUndo[sectionKey].length > 20) ent.sectionUndo[sectionKey].shift();
    const entry = stack.pop();
    ent[sectionKey] = JSON.parse(JSON.stringify(entry.data || []));
    appendLog(`Redid ${sectionKey} change`, ent.id);
    saveAndRefresh();
  }

  function sectionRedoTitle(ent, sectionKey, emptyMsg = 'Nothing to redo.') {
    ensureSectionUndoStore(ent);
    const stack = ent.sectionRedo[sectionKey];
    if (!stack || !stack.length) return emptyMsg;
    const last = stack[stack.length - 1];
    const when = last.time ? new Date(last.time).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '';
    return `Redo: ${last.label || 'last change'}${when ? ` (${when})` : ''}`;
  }

  function undoLastDestructiveAction() {
    const camp = currentCampaign();
    if (!camp || !Array.isArray(camp.undoStack) || !camp.undoStack.length) {
      showToast('Nothing to undo.', 'warn');
      return;
    }
    if (!Array.isArray(camp.redoStack)) camp.redoStack = [];
    const currentSnapshot = JSON.parse(JSON.stringify(camp));
    currentSnapshot.undoStack = [];
    currentSnapshot.redoStack = [];
    const entry = camp.undoStack.pop();
    camp.redoStack.push({
      id: generateId('redo'),
      time: new Date().toISOString(),
      label: entry.label || 'last change',
      target: entry.target || '',
      targetLabel: entry.targetLabel || '',
      snapshot: currentSnapshot
    });
    if (camp.redoStack.length > 20) camp.redoStack.shift();
    const restored = JSON.parse(JSON.stringify(entry.snapshot || {}));
    restored.undoStack = camp.undoStack.slice();
    restored.redoStack = camp.redoStack.slice();
    isApplyingUndo = true;
    state.campaigns[state.currentCampaignId] = restored;
    isApplyingUndo = false;
    saveAndRefresh();
    renderLog();
    renderSessionPrep();
    updateUndoButtonState();
    const targetInfo = entry.targetLabel ? ` — ${entry.targetLabel}` : '';
    showToast(`Undid: ${entry.label || 'last change'}${targetInfo}`, 'info');
  }

  function redoLastDestructiveAction() {
    const camp = currentCampaign();
    if (!camp || !Array.isArray(camp.redoStack) || !camp.redoStack.length) {
      showToast('Nothing to redo.', 'warn');
      return;
    }
    if (!Array.isArray(camp.undoStack)) camp.undoStack = [];
    const currentSnapshot = JSON.parse(JSON.stringify(camp));
    currentSnapshot.undoStack = [];
    currentSnapshot.redoStack = [];
    const entry = camp.redoStack.pop();
    camp.undoStack.push({
      id: generateId('undo'),
      time: new Date().toISOString(),
      label: entry.label || 'redo action',
      target: entry.target || '',
      targetLabel: entry.targetLabel || '',
      snapshot: currentSnapshot
    });
    if (camp.undoStack.length > 20) camp.undoStack.shift();
    const restored = JSON.parse(JSON.stringify(entry.snapshot || {}));
    restored.undoStack = camp.undoStack.slice();
    restored.redoStack = camp.redoStack.slice();
    isApplyingUndo = true;
    state.campaigns[state.currentCampaignId] = restored;
    isApplyingUndo = false;
    saveAndRefresh();
    renderLog();
    renderSessionPrep();
    updateUndoButtonState();
    const targetInfo = entry.targetLabel ? ` — ${entry.targetLabel}` : '';
    showToast(`Redid: ${entry.label || 'last change'}${targetInfo}`, 'info');
  }

  function deleteEntity(entityId, options = {}) {
    const camp = currentCampaign();
    const ent = camp.entities[entityId];
    if (!ent) return false;

    const actionLabel = options.actionLabel || ('Deleted ' + ent.type.toUpperCase());
    appendLog(actionLabel, entityId);

    // Remove relationships that touch this entity.
    Object.keys(camp.relationships).forEach((rid) => {
      const rel = camp.relationships[rid];
      if (rel.source === entityId || rel.target === entityId) {
        delete camp.relationships[rid];
      }
    });

    // Remove references from all entities.
    Object.values(camp.entities).forEach((other) => {
      if (other.id === entityId) return;
      if (Array.isArray(other.members)) {
        other.members = other.members.filter(id => id !== entityId);
      }
      if (other.affiliation === entityId) {
        other.affiliation = '';
      }
    });

    delete camp.entities[entityId];
    delete camp.positions[entityId];
    if (camp.playerOwnedPcId === entityId) camp.playerOwnedPcId = null;
    if (state.selectedEntityId === entityId) state.selectedEntityId = null;
    if (graphSelectedNodeId === entityId) graphSelectedNodeId = null;
    if (state.selectedRelId && !camp.relationships[state.selectedRelId]) state.selectedRelId = null;
    return true;
  }

  function deleteRelationship(relId, options = {}) {
    const camp = currentCampaign();
    const rel = camp.relationships[relId];
    if (!rel) return false;
    const actionLabel = options.actionLabel || 'Deleted relationship';
    delete camp.relationships[relId];
    appendLog(actionLabel, relId);
    if (state.selectedRelId === relId) state.selectedRelId = null;
    return true;
  }

  async function confirmAndDeleteEntity(entityId, actionLabel) {
    const camp = currentCampaign();
    const ent = camp.entities[entityId];
    if (!ent) return;
    const ok = await askConfirm(
      `Delete ${ent.type.toUpperCase()} "${entityLabel(ent)}"? This also removes linked relationships.`,
      'Delete Entity'
    );
    if (!ok) return;
    captureUndoSnapshot(`Delete ${ent.type.toUpperCase()} "${entityLabel(ent)}"`, ent.id);
    if (deleteEntity(entityId, { actionLabel })) {
      saveAndRefresh();
    }
  }

  async function confirmAndDeleteRelationship(relId, actionLabel) {
    const camp = currentCampaign();
    const rel = camp.relationships[relId];
    if (!rel) return;
    const ok = await askConfirm('Are you sure you want to delete this relationship?', 'Delete Relationship');
    if (!ok) return;
    captureUndoSnapshot(`Delete relationship ${rel.type || ''}`.trim(), rel.id);
    if (deleteRelationship(relId, { actionLabel })) {
      saveAndRefresh();
      const inspector = document.getElementById('inspector');
      if (inspector) inspector.classList.add('hidden');
    }
  }

  /**
   * Toggle GM mode. Update state, re-render all views to show or hide
   * secret information, and persist the setting.
   */
  function toggleGMMode(forceValue) {
    state.gmMode = forceValue !== undefined ? forceValue : !state.gmMode;
    const toggle = document.getElementById('gm-toggle');
    if (toggle) toggle.checked = state.gmMode;
    currentCampaign().gmMode = state.gmMode;
    applyModeClasses();
    renderEntityLists();
    renderSheetView();
    updateGraph();
    renderLog();
    renderMessages();
    saveCampaigns();
  }

  /**
   * Toggle dark mode. Applies a class on the body and stores the setting.
   */
  function toggleDarkMode(forceValue) {
    state.darkMode = forceValue !== undefined ? forceValue : !state.darkMode;
    const body = document.body;
    if (state.darkMode) {
      body.classList.add('dark');
    } else {
      body.classList.remove('dark');
    }
    // Update icon: sun in dark mode (to switch to light), moon in light mode
    const darkBtn = document.getElementById('dark-mode-btn');
    if (darkBtn) {
      const icon = darkBtn.querySelector('.material-icons');
      if (icon) icon.textContent = state.darkMode ? 'light_mode' : 'dark_mode';
      darkBtn.title = state.darkMode ? 'Switch to light mode' : 'Switch to dark mode';
    }
    currentCampaign().darkMode = state.darkMode;
    saveCampaigns();
  }

  /**
   * Render the campaign name in the top navigation bar.
   */
  function renderCampaignName() {
    const span = document.getElementById('campaign-name');
    span.textContent = currentCampaign().name;
  }

  /**
   * Show a quick-view popover next to a sidebar list item.
   */
  let _popoverTimer = null;
  function showSidebarPopover(ent, anchorEl) {
    clearTimeout(_popoverTimer);
    _popoverTimer = setTimeout(() => {
      let pop = document.getElementById('sidebar-popover');
      if (!pop) {
        pop = document.createElement('div');
        pop.id = 'sidebar-popover';
        pop.className = 'sidebar-popover';
        document.body.appendChild(pop);
      }
      pop.innerHTML = '';
      const camp = currentCampaign();

      const nameEl = document.createElement('div');
      nameEl.className = 'sp-name';
      nameEl.textContent = entityLabel(ent);
      pop.appendChild(nameEl);

      if (ent.type === 'pc') {
        const lines = [];
        if (ent.class) lines.push(ent.class + (ent.durance ? ' / ' + ent.durance : ''));
        // Stress summary
        const tracks = ['blood','mind','silver','shadow','reputation'];
        const stressed = tracks.filter(t => ent.stressFilled && ent.stressFilled[t] && ent.stressFilled[t].length > 0);
        if (stressed.length) {
          lines.push('Stress: ' + stressed.map(t => t.charAt(0).toUpperCase() + ':' + ent.stressFilled[t].length + '/' + getTrackTotalSlots(ent, t)).join(' '));
        }
        if (ent.fallout && ent.fallout.filter(f => !f.resolved).length > 0)
          lines.push(ent.fallout.filter(f => !f.resolved).length + ' active fallout');
        lines.forEach(l => { const d = document.createElement('div'); d.className = 'sp-line'; d.textContent = l; pop.appendChild(d); });
      } else if (ent.type === 'npc') {
        const lines = [];
        if (ent.role) lines.push(ent.role);
        if (ent.disposition) lines.push('Disposition: ' + ent.disposition);
        if (ent.threatLevel) lines.push('Threat: ' + ent.threatLevel);
        if (ent.affiliation && camp.entities[ent.affiliation]) lines.push('Org: ' + entityLabel(camp.entities[ent.affiliation]));
        lines.forEach(l => { const d = document.createElement('div'); d.className = 'sp-line'; d.textContent = l; pop.appendChild(d); });
      } else if (ent.type === 'org') {
        const lines = [];
        if (ent.reach) lines.push('Reach: ' + ent.reach);
        if (ent.ministryRelation) lines.push('Ministry: ' + ent.ministryRelation);
        if (ent.members && ent.members.length) lines.push('Members: ' + ent.members.length);
        lines.forEach(l => { const d = document.createElement('div'); d.className = 'sp-line'; d.textContent = l; pop.appendChild(d); });
      }

      const rect = anchorEl.getBoundingClientRect();
      const sidebarRect = document.getElementById('sidebar').getBoundingClientRect();
      pop.style.left = (sidebarRect.right + 6) + 'px';
      pop.style.top = Math.min(rect.top, window.innerHeight - 160) + 'px';
      pop.classList.add('visible');
    }, 350);
  }
  function hideSidebarPopover() {
    clearTimeout(_popoverTimer);
    const pop = document.getElementById('sidebar-popover');
    if (pop) pop.classList.remove('visible');
  }

  /**
   * Render the lists of PCs, NPCs, and Orgs in the sidebar. Each list item
   * becomes clickable to select an entity. Omits GM-only entities in player
   * mode. Updates count badges and adds coloured type dots inline.
   */
  function renderEntityLists() {
    const camp = currentCampaign();
    const searchTerm = document.getElementById('search-input').value.trim().toLowerCase();
    const sortMode = ['manual', 'name', 'pinned'].includes(camp.entitySort) ? camp.entitySort : 'manual';
    const sortSelect = document.getElementById('entity-sort-select');
    if (sortSelect && sortSelect.value !== sortMode) sortSelect.value = sortMode;
    const pinnedOnly = !!camp.entityPinnedOnly;
    const pinOnlyChk = document.getElementById('entity-pin-only');
    if (pinOnlyChk) pinOnlyChk.checked = pinnedOnly;
    const lists = {
      pc:  document.getElementById('pc-list'),
      npc: document.getElementById('npc-list'),
      org: document.getElementById('org-list')
    };

    // Clear lists
    Object.values(lists).forEach(l => { l.innerHTML = ''; });

    // Populate
    const entities = Object.values(camp.entities).filter((ent) => {
      if (!state.gmMode && ent.gmOnly) return;
      if (pinnedOnly && !ent.pinned) return;
      const label = entityLabel(ent).toLowerCase();
      if (searchTerm && !label.includes(searchTerm)) return;
      return true;
    });
    entities.sort((a, b) => {
      if (sortMode === 'name') {
        return entityLabel(a).localeCompare(entityLabel(b), undefined, { sensitivity: 'base' });
      }
      if (sortMode === 'pinned') {
        const aPinned = !!a.pinned;
        const bPinned = !!b.pinned;
        if (aPinned !== bPinned) return aPinned ? -1 : 1;
        return entityLabel(a).localeCompare(entityLabel(b), undefined, { sensitivity: 'base' });
      }
      return 0;
    });
    entities.forEach(ent => {

      const li = document.createElement('li');
      li.dataset.id = ent.id;
      if (state.selectedEntityId === ent.id) li.classList.add('active');

      // Coloured dot
      const dot = document.createElement('span');
      dot.className = 'entity-dot';
      dot.style.background = ent.type === 'pc'  ? 'var(--pc-color)'
                           : ent.type === 'npc' ? 'var(--npc-color)'
                           :                      'var(--org-color)';
      li.appendChild(dot);

      const name = document.createElement('span');
      name.className = 'entity-name';
      name.textContent = entityLabel(ent) || '(Unnamed)';
      name.addEventListener('dblclick', (e) => {
        e.preventDefault();
        e.stopPropagation();
        state.selectedRelId = null;
        selectEntity(ent.id);
        const sheetsTab = document.querySelector('.tab-link[data-tab="sheets-view"]');
        if (sheetsTab) sheetsTab.click();
      });
      li.appendChild(name);

      const meta = document.createElement('span');
      meta.className = 'entity-meta';

      if (state.gmMode) {
        const pinBtn = document.createElement('button');
        pinBtn.type = 'button';
        pinBtn.className = 'entity-pin-btn' + (ent.pinned ? ' active' : '');
        pinBtn.title = ent.pinned ? 'Unpin' : 'Pin';
        pinBtn.setAttribute('aria-label', ent.pinned ? 'Unpin entity' : 'Pin entity');
        pinBtn.textContent = ent.pinned ? '★' : '☆';
        pinBtn.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          ent.pinned = !ent.pinned;
          appendLog(ent.pinned ? 'Pinned entity' : 'Unpinned entity', ent.id);
          saveCampaigns();
          renderEntityLists();
          showToast(ent.pinned ? 'Pinned entity.' : 'Unpinned entity.', 'info');
        });
        meta.appendChild(pinBtn);
      }

      // GM-only badge
      if (ent.gmOnly && state.gmMode) {
        const b = document.createElement('span');
        b.className = 'gm-badge';
        b.textContent = 'GM';
        meta.appendChild(b);
      }

      // Pending approval badge (visible to both GM and the submitting player)
      if (ent.pendingApproval) {
        const b = document.createElement('span');
        b.className = 'pending-badge';
        b.textContent = 'Pending';
        meta.appendChild(b);
      }
      if (meta.children.length) li.appendChild(meta);

      li.addEventListener('click', () => {
        state.selectedRelId = null;
        selectEntity(ent.id);
      });
      li.addEventListener('dblclick', () => {
        state.selectedRelId = null;
        selectEntity(ent.id);
        const sheetsTab = document.querySelector('.tab-link[data-tab=\"sheets-view\"]');
        if (sheetsTab) sheetsTab.click();
      });
      // Hover popover
      li.addEventListener('mouseenter', e => showSidebarPopover(ent, li));
      li.addEventListener('mouseleave', () => hideSidebarPopover());

      lists[ent.type].appendChild(li);
    });

    // Update count badges
    ['pc', 'npc', 'org'].forEach(t => {
      const el = document.getElementById(t + '-count');
      if (el) el.textContent = lists[t].children.length;
    });
    ensureAccessibilityLabels(document.getElementById('sidebar'));
  }

  /**
   * Generate a display label for an entity based on its type and fields.
   */
  function entityLabel(ent) {
    if (ent.type === 'pc') {
      const name = [ent.firstName, ent.lastName].filter(x => x).join(' ').trim();
      return name || ent.name || 'Unnamed PC';
    }
    return ent.name || (ent.type.toUpperCase());
  }

  /**
   * Select an entity by ID. Renders its sheet in the sheet view and clears
   * any selected relationship.
   */
  function selectEntity(id) {
    state.selectedEntityId = id;
    state.selectedRelId = null;
    graphSelectedNodeId = id;
    renderEntityLists();
    renderSheetView();
    // Optionally open inspector panel automatically for small screens
  }

  function visibleEntityOrder() {
    const order = [];
    ['pc-list', 'npc-list', 'org-list'].forEach((listId) => {
      const list = document.getElementById(listId);
      if (!list) return;
      Array.from(list.querySelectorAll('li[data-id]')).forEach((li) => {
        const id = li.dataset.id;
        if (id) order.push(id);
      });
    });
    return order;
  }

  function selectAdjacentEntity(step = 1) {
    const ids = visibleEntityOrder();
    if (!ids.length) return;
    const current = state.selectedEntityId;
    const idx = ids.indexOf(current);
    const nextIdx = idx === -1
      ? 0
      : (idx + step + ids.length) % ids.length;
    selectEntity(ids[nextIdx]);
    const sheetsTab = document.querySelector('.tab-link[data-tab="sheets-view"]');
    if (sheetsTab) sheetsTab.click();
  }

  function selectAdjacentTab(step = 1) {
    const tabs = Array.from(document.querySelectorAll('.tab-link'))
      .filter((btn) => btn && btn.offsetParent !== null);
    if (!tabs.length) return;
    const current = document.querySelector('.tab-link.active');
    const idx = Math.max(0, tabs.indexOf(current));
    const nextIdx = (idx + step + tabs.length) % tabs.length;
    const btn = tabs[nextIdx];
    if (btn) btn.click();
  }

  function quickAddTaskToSelectedPC() {
    const camp = currentCampaign();
    const ent = state.selectedEntityId ? camp.entities[state.selectedEntityId] : null;
    if (!ent || ent.type !== 'pc') {
      showToast('Select a PC sheet first to add a task.', 'warn');
      return;
    }
    const isOwnPC = !state.gmMode && camp.playerOwnedPcId === ent.id && camp.allowPlayerEditing;
    const canEditPC = state.gmMode || isOwnPC;
    if (!canEditPC) {
      showToast('You cannot edit this character.', 'warn');
      return;
    }
    if (!Array.isArray(ent.tasks)) ent.tasks = [];
    ensureSectionUndoStore(ent);
    pushSectionUndo(ent, 'tasks', 'Quick add task');
    ent.tasks.push({ id: generateId('task'), title: '', status: 'To Do', priority: 'Normal', dueDate: '', notes: '' });
    appendLog('Added task', ent.id);
    saveAndRefresh();
  }

  /**
   * Select a relationship by ID. Opens the inspector for editing the
   * relationship properties.
   */
  function selectRelationship(id) {
    state.selectedRelId = id;
    state.selectedEntityId = null;
    renderEntityLists();
    renderSheetView();
    renderInspectorForRelationship();
  }

  /**
   * Render the appropriate sheet view based on the selected entity or
   * relationship. If nothing is selected, show a placeholder.
   */
  function renderSheetView() {
    const sheetContainer = document.getElementById('sheet-container');
    const placeholder = document.getElementById('sheet-placeholder');
    sheetContainer.innerHTML = '';
    if (state.selectedEntityId) {
      const ent = currentCampaign().entities[state.selectedEntityId];
      if (!ent) return;
      placeholder.style.display = 'none';
      sheetContainer.style.display = 'block';
      if (ent.type === 'pc') {
        renderPCSheet(ent);
      } else if (ent.type === 'npc') {
        renderNPCSheet(ent);
      } else if (ent.type === 'org') {
        renderOrgSheet(ent);
      }
    } else {
      sheetContainer.style.display = 'none';
      placeholder.style.display = 'block';
    }
  }

  // -----------------------------------------------------------------------
  // HELP TOOLTIP TEXT for PC sheet sections
  // -----------------------------------------------------------------------
  const SECTION_HELP = {
    stress: 'Each resistance track represents a kind of pressure. When you take stress, check Fallout using your total stress across tracks (extra slots above 10 per track do not add to Fallout total).',
    fallout: 'After stress is taken, roll D10 against total stress. If the roll is lower than total stress, Fallout triggers: Minor (2-4), Moderate (5-8), Severe (9+).',
    skills: 'Skills are either known or Mastered (shown as ★). Mastery and relevant Domains add dice; keep the highest result.',
    domains: 'Domains represent social or professional contexts. If a domain applies, you roll an extra die and keep the highest.',
    resistances: 'Resistances are stress tracks. A higher value means more stress capacity before consequences, and extra slots beyond 10 do not increase Fallout total.',
    bonds: 'Bonds are your connections to people and groups. Individual bonds are personal; Street bonds connect you to a community; Organisation bonds tie you to a faction.',
    refresh: 'Your core ability recharges when you meet this condition. Mark as Refreshed to track it.',
    advances: 'Spend Advance Points earned through play to unlock new abilities. Low advances cost 1 AP, Medium 2 AP, High 3 AP. You must have 2 Low advances before taking Medium, and 2 Medium before High.',
    inventory: 'Your carried equipment. Weapons deal stress on a hit. Armour reduces incoming stress by its resistance value.',
    tasks: 'Ongoing objectives and quests. Track status and priority here.',
    history: 'Recent changes and actions recorded for this character sheet.',
    ministry: 'The Ministry of Our Lady of the Dual Chain — the Aelfir rulers of Spire — may be aware of your activities. Higher attention means more scrutiny and danger.'
  };

  // Create a collapsible inspector-section with a help tooltip
  function makeSection(id, title, helpKey, pc, forceOpen = false) {
    const camp = currentCampaign();
    if (!camp.sectionCollapse) camp.sectionCollapse = {};
    const collapseKey = (pc ? pc.id + ':' : '') + id;
    // Default: open unless previously collapsed
    const isOpen = camp.sectionCollapse[collapseKey] !== false || forceOpen;

    const sec = document.createElement('div');
    sec.className = 'inspector-section collapsible-section' + (isOpen ? ' open' : ' collapsed');
    sec.dataset.collapseKey = collapseKey;
    sec.dataset.sectionId = id;

    const header = document.createElement('div');
    header.className = 'section-header';

    const arrow = document.createElement('span');
    arrow.className = 'section-arrow';
    arrow.textContent = isOpen ? '▾' : '▸';

    const titleEl = document.createElement('span');
    titleEl.className = 'section-title';
    titleEl.textContent = title;

    header.appendChild(arrow);
    header.appendChild(titleEl);

    if (helpKey && SECTION_HELP[helpKey]) {
      const help = document.createElement('span');
      help.className = 'section-help';
      help.textContent = 'ⓘ';
      help.title = SECTION_HELP[helpKey];
      help.addEventListener('click', e => {
        e.stopPropagation();
        showToast(SECTION_HELP[helpKey], 'info');
      });
      header.appendChild(help);
    }

    const body = document.createElement('div');
    body.className = 'section-body';

    header.addEventListener('click', () => {
      const open = sec.classList.toggle('open');
      sec.classList.toggle('collapsed', !open);
      arrow.textContent = open ? '▾' : '▸';
      camp.sectionCollapse[collapseKey] = open;
    });

    sec.appendChild(header);
    sec.appendChild(body);
    return { sec, body };
  }

  /**
   * Render a sheet for a PC. Generates dynamic fields for identity,
   * stress, fallout, skills, domains, resistances, bonds, inventory and
   * tasks. Events are attached inline to support editing in place.
   */
  function renderPCSheet(pc) {
    const container = document.getElementById('sheet-container');
    // On initial render, ensure class and durance effects are applied if
    // necessary. If no items with the appropriate source exist, apply
    // effects once. This helps maintain backwards compatibility when
    // loading older campaigns that did not track these modifications.
    if (pc.class && !pc.skills.some(s => s.source && s.source.startsWith('class:'))) {
      applyClassEffects(pc, pc.class);
    }
    if (pc.durance && !pc.skills.some(s => s.source && s.source.startsWith('durance:'))) {
      applyDuranceEffects(pc, pc.durance);
    }
    // Consolidate any duplicate resistance names (e.g. both class and durance
    // contribute to the same track) into a single merged row.
    consolidateResistances(pc);
    // Migrate missing fields
    if (!Array.isArray(pc.bonds)) pc.bonds = [];
    ensureSectionUndoStore(pc);
    if (pc.refreshed === undefined) pc.refreshed = false;
    if (pc.advancePoints === undefined) pc.advancePoints = 0;
    const campPerm = currentCampaign();
    const isOwnPC = !state.gmMode && pc.type === 'pc' && campPerm.playerOwnedPcId === pc.id && campPerm.allowPlayerEditing;
    const canEditPC = state.gmMode || isOwnPC;
    container.innerHTML = '';
    // Title row: name + action buttons
    const title = document.createElement('div');
    title.className = 'sheet-title';
    const name = escapeHtml(entityLabel(pc));
    const classSuffix = pc.class ? ` <span class="sheet-class-suffix">— ${escapeHtml(pc.class)}</span>` : '';
    const isMyCharacter = !state.gmMode && currentCampaign().playerOwnedPcId === pc.id;
    const myTag = isMyCharacter ? ' <span class="sheet-class-suffix" style="margin-left:8px;color:var(--pl-accent-hi)">My Character</span>' : '';
    title.innerHTML = `<h2>${name}${classSuffix}${myTag}</h2>`;
    const titleBtns = document.createElement('div');
    titleBtns.style.display = 'flex';
    titleBtns.style.gap = '6px';
    // Duplicate button (GM only)
    if (state.gmMode) {
      const dupBtn = document.createElement('button');
      dupBtn.className = 'print-btn';
      dupBtn.title = 'Duplicate entity';
      dupBtn.textContent = 'Duplicate';
      dupBtn.addEventListener('click', () => {
        const camp = currentCampaign();
        const clone = JSON.parse(JSON.stringify(pc));
        clone.id = generateId('pc');
        clone.name = (clone.name || 'Copy') + ' (copy)';
        clone.firstName = clone.firstName ? clone.firstName + ' (copy)' : 'Copy';
        camp.entities[clone.id] = clone;
        appendLog('Duplicated PC', clone.id);
        saveAndRefresh();
        selectEntity(clone.id);
      });
      titleBtns.appendChild(dupBtn);

      const delBtn = document.createElement('button');
      delBtn.className = 'print-btn print-btn-danger';
      delBtn.title = 'Delete entity';
      delBtn.textContent = 'Delete';
      delBtn.addEventListener('click', async () => {
        await confirmAndDeleteEntity(pc.id);
      });
      titleBtns.appendChild(delBtn);
    }
    const printBtn = document.createElement('button');
    printBtn.className = 'print-btn';
    printBtn.textContent = 'Print';
    printBtn.addEventListener('click', () => printEntitySheet(pc.id));
    titleBtns.appendChild(printBtn);
    const cardBtn = document.createElement('button');
    cardBtn.className = 'print-btn';
    cardBtn.textContent = 'Export ID Card';
    cardBtn.title = 'Printable portrait and identity card';
    cardBtn.addEventListener('click', () => exportPCIdentityCard(pc.id));
    titleBtns.appendChild(cardBtn);
    if (state.gmMode) {
      const printPlayerBtn = document.createElement('button');
      printPlayerBtn.className = 'print-btn';
      printPlayerBtn.textContent = 'Print Player Copy';
      printPlayerBtn.title = 'One-click player-safe print';
      printPlayerBtn.addEventListener('click', () => printEntitySheet(pc.id, { forcePlayerSafe: true }));
      titleBtns.appendChild(printPlayerBtn);
    }
    title.appendChild(titleBtns);
    container.appendChild(title);
    const sectionJumpNav = document.createElement('div');
    sectionJumpNav.className = 'pc-section-nav';
    [
      ['identity', 'Identity'],
      ['stress', 'Stress'],
      ['fallout', 'Fallout'],
      ['skills', 'Skills'],
      ['domains', 'Domains'],
      ['resistances', 'Resist'],
      ['bonds', 'Bonds'],
      ['inventory', 'Inventory'],
      ['tasks', 'Tasks'],
      ['history', 'History']
    ].forEach(([secId, label]) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = label;
      btn.addEventListener('click', () => {
        const target = container.querySelector(`[data-section-id="${secId}"]`);
        if (target && target.scrollIntoView) {
          target.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
      });
      sectionJumpNav.appendChild(btn);
    });
    container.appendChild(sectionJumpNav);

    // Two-column layout
    const twoCol = document.createElement('div');
    twoCol.className = 'sheet-two-col';
    const leftCol = document.createElement('div');
    leftCol.className = 'sheet-col sheet-col-left';
    const rightCol = document.createElement('div');
    rightCol.className = 'sheet-col sheet-col-right';
    twoCol.appendChild(leftCol);
    twoCol.appendChild(rightCol);

    // Identity section (compact width above columns)
    const identity = document.createElement('div');
    identity.className = 'inspector-section pc-identity-section';
    identity.dataset.sectionId = 'identity';
    identity.innerHTML = `<h3>Identity</h3>`;
    identity.appendChild(createInputField(pc, 'firstName', 'First name'));
    identity.appendChild(createInputField(pc, 'lastName', 'Last name'));
    identity.appendChild(createInputField(pc, 'pronouns', 'Pronouns (optional)'));
    identity.appendChild(createPortraitField(pc, 'Character Portrait'));
    // Class dropdown
    const classField = document.createElement('div');
    classField.className = 'inspector-field';
    const classLabel = document.createElement('label');
    classLabel.textContent = 'Class';
    const classSelect = document.createElement('select');
    const classes = ['','Azurite','Blood Witch','Bound','Carrion-Priest','Firebrand','Idol','Knight','Lahjan','Masked','Midwife','Vermissian Sage'];
    classes.forEach(cls => {
      const opt = document.createElement('option');
      opt.value = cls;
      opt.textContent = cls || 'Select class';
      if (pc.class === cls) opt.selected = true;
      classSelect.appendChild(opt);
    });
    classSelect.addEventListener('change', e => {
      const newClass = e.target.value;
      // Apply class effects before updating property to ensure removal of
      // previous modifications. Apply new effects afterwards.
      applyClassEffects(pc, newClass);
      pc.class = newClass;
      appendLog('Changed class', pc.id);
      saveAndRefresh();
    });
    classField.appendChild(classLabel);
    classField.appendChild(classSelect);
    identity.appendChild(classField);
    // Durance dropdown
    const durField = document.createElement('div');
    durField.className = 'inspector-field';
    const durLabel = document.createElement('label');
    durLabel.textContent = 'Durance';
    const durSelect = document.createElement('select');
    const durOpt = document.createElement('option');
    durOpt.value = '';
    durOpt.textContent = 'Select durance';
    durSelect.appendChild(durOpt);
    DURANCE_OPTIONS.forEach(dur => {
      const opt = document.createElement('option');
      opt.value = dur;
      opt.textContent = dur;
      if (pc.durance === dur) opt.selected = true;
      durSelect.appendChild(opt);
    });
    durSelect.addEventListener('change', e => {
      const newDur = e.target.value;
      applyDuranceEffects(pc, newDur);
      pc.durance = newDur;
      appendLog('Changed durance', pc.id);
      saveAndRefresh();
    });
    durField.appendChild(durLabel);
    durField.appendChild(durSelect);
    identity.appendChild(durField);
    container.appendChild(twoCol);
    leftCol.appendChild(identity);
    // LEFT COLUMN: Stress, Fallout, Bonds, Inventory, Tasks
    // RIGHT COLUMN: Skills, Domains, Resistances, Class Features
    // Stress section
    const stressSection = document.createElement('div');
    stressSection.className = 'inspector-section';
    stressSection.dataset.sectionId = 'stress';
    stressSection.innerHTML = '<h3>Stress</h3>';

    // Ensure stressSlots and stressFilled exist (migrate old saves gracefully)
    if (!pc.stressSlots) {
      pc.stressSlots = { blood: 10, mind: 10, silver: 10, shadow: 10, reputation: 10 };
    }
    if (!pc.stressFilled || typeof pc.stressFilled.blood !== 'object') {
      pc.stressFilled = {};
      ['blood','mind','silver','shadow','reputation'].forEach(track => {
        const old = (pc.stress && typeof pc.stress[track] === 'number') ? pc.stress[track] : 0;
        pc.stressFilled[track] = Array.from({ length: old }, (_, i) => i);
      });
    }

    ['blood','mind','silver','shadow','reputation'].forEach(track => {
      const baseSlots = getTrackBaseSlots(pc, track);
      const freeSlots = getTrackFreeSlots(pc, track);
      const slots = getTrackTotalSlots(pc, track);
      const filled = pc.stressFilled[track] || [];

      const row = document.createElement('div');
      row.className = `stress-row stress-${track}`;

      // Label
      const label = document.createElement('span');
      label.className = 'stress-label';
      label.textContent = track.charAt(0).toUpperCase() + track.slice(1);
      row.appendChild(label);

      // Slot controls (+/-)
      const controls = document.createElement('div');
      controls.className = 'stress-controls';

      const minusBtn = document.createElement('button');
      minusBtn.textContent = '−';
      minusBtn.title = 'Remove stress slot';
      minusBtn.addEventListener('click', () => {
        if (getTrackBaseSlots(pc, track) <= 1) return;
        pc.stressSlots[track]--;
        // Truncate stress to new cap without triggering fallout.
        const current = (pc.stressFilled[track] || []).length;
        setPCStressLevel(pc, track, Math.min(current, getTrackTotalSlots(pc, track)), { triggerFallout: false });
        appendLog(`Reduced ${track} stress slots`, pc.id);
        saveAndRefresh();
      });

      const slotCount = document.createElement('span');
      slotCount.className = 'stress-slot-count';
      slotCount.textContent = freeSlots > 0 ? `${baseSlots} +${freeSlots} free` : String(baseSlots);

      const plusBtn = document.createElement('button');
      plusBtn.textContent = '+';
      plusBtn.title = 'Add stress slot';
      plusBtn.addEventListener('click', () => {
        if (getTrackBaseSlots(pc, track) >= 20) return;
        pc.stressSlots[track]++;
        appendLog(`Increased ${track} stress slots`, pc.id);
        saveAndRefresh();
      });

      controls.appendChild(minusBtn);
      controls.appendChild(slotCount);
      controls.appendChild(plusBtn);
      row.appendChild(controls);

      // Numeric stress counter
      const stressCounter = document.createElement('span');
      stressCounter.className = 'stress-num-counter';
      const counted = getTrackFalloutCountedStress(pc, track);
      stressCounter.textContent = freeSlots > 0
        ? `${filled.length}/${slots} (${counted} counted)`
        : `${filled.length}/${slots}`;
      row.appendChild(stressCounter);

      // Pips — sequential fill
      const pipsContainer = document.createElement('div');
      pipsContainer.className = 'stress-pips';
      for (let i = 0; i < slots; i++) {
        const pip = document.createElement('div');
        const idx = i;
        const filledCount = filled.filter(x => x <= idx).length; // used only for initial render
        pip.className = 'pip' + (filled.includes(i) ? ' active' : '');
        pip.title = `Stress ${i + 1}`;
        makeInteractivePip(pip, `${track} stress ${i + 1}`);
        pip.addEventListener('click', () => {
          const arr = pc.stressFilled[track];
          // Determine current fill level = highest active index + 1
          const currentMax = arr.length > 0 ? Math.max(...arr) : -1;
          if (idx <= currentMax) {
            // Clicked at or below current level — clear back to idx-1
            setPCStressLevel(pc, track, idx);
          } else {
            // Clicked above current level — fill up to idx
            setPCStressLevel(pc, track, idx + 1);
          }
          saveAndRefresh();
        });
        pipsContainer.appendChild(pip);
      }
      row.appendChild(pipsContainer);
      stressSection.appendChild(row);
    });
    // Bulk clear stress button
    if (state.gmMode) {
      const clearBtn = document.createElement('button');
      clearBtn.textContent = 'Clear All Stress';
      clearBtn.style.marginTop = '6px';
      clearBtn.style.fontSize = '0.78rem';
      clearBtn.addEventListener('click', () => {
        ['blood','mind','silver','shadow','reputation'].forEach(t => setPCStressLevel(pc, t, 0, { triggerFallout: false }));
        appendLog('Cleared all stress', pc.id);
        saveAndRefresh();
      });
      stressSection.appendChild(clearBtn);
    }
    const applyStressBtn = document.createElement('button');
    applyStressBtn.textContent = 'Apply Stress…';
    applyStressBtn.style.marginTop = '6px';
    applyStressBtn.style.marginLeft = '6px';
    applyStressBtn.style.fontSize = '0.78rem';
    applyStressBtn.disabled = !canEditPC;
    applyStressBtn.addEventListener('click', () => {
      openStressApplicationModal(pc);
    });
    stressSection.appendChild(applyStressBtn);
    leftCol.appendChild(stressSection);

    // Fallout section → left col
    const { sec: falloutSec, body: falloutBody } = makeSection('fallout', 'Fallout', 'fallout', pc);
    // Count active vs resolved
    const activeFallout = pc.fallout.filter(f => !f.resolved);
    const resolvedFallout = pc.fallout.filter(f => f.resolved);
    if (activeFallout.length) {
      const activeHeader = document.createElement('div');
      activeHeader.className = 'fallout-subsection-header';
      activeHeader.textContent = 'Active';
      falloutBody.appendChild(activeHeader);
      activeFallout.forEach(f => falloutBody.appendChild(createFalloutRow(pc, f)));
    }
    if (resolvedFallout.length) {
      const resolvedHeader = document.createElement('div');
      resolvedHeader.className = 'fallout-subsection-header resolved';
      resolvedHeader.textContent = 'Resolved';
      falloutBody.appendChild(resolvedHeader);
      resolvedFallout.forEach(f => falloutBody.appendChild(createFalloutRow(pc, f)));
    }
    if (!activeFallout.length && !resolvedFallout.length) {
      const empty = document.createElement('p');
      empty.className = 'text-muted';
      empty.style.fontSize = '0.82rem';
      empty.textContent = 'No fallout recorded. Add fallout when stress checks trigger consequences.';
      falloutBody.appendChild(empty);
    }
    const addFalloutBtn = document.createElement('button');
    addFalloutBtn.textContent = '+ Add Fallout';
    addFalloutBtn.addEventListener('click', () => {
      pc.fallout.push({ id: generateId('fallout'), type: 'Blood', severity: 'Minor', name: '', description: '', resolved: false, timestamp: new Date().toISOString() });
      appendLog('Added fallout', pc.id);
      saveAndRefresh();
    });
    falloutBody.appendChild(addFalloutBtn);
    leftCol.appendChild(falloutSec);
    // Resistances → right col (with stress context bars)
    const { sec: resSec, body: resBody } = makeSection('resistances', 'Resistances', 'resistances', pc);
    pc.resistances.forEach(r => {
      const resRow = document.createElement('div');
      resRow.className = 'resistance-display-row';
      const nameSpan = document.createElement('span');
      nameSpan.className = 'res-name';
      nameSpan.textContent = r.name;
      const valBadge = document.createElement('span');
      valBadge.className = 'res-value-badge';
      valBadge.textContent = '+' + (r.value || 0) + ' free';
      const rollBtn = document.createElement('button');
      rollBtn.className = 'res-roll-btn';
      rollBtn.textContent = 'Roll';
      rollBtn.title = `Roll D10 against ${r.name} resistance`;
      rollBtn.addEventListener('click', () => {
        openDiceRollerForResistance(r.name, parseInt(r.value || 0, 10), pc);
      });
      // Stress bar for this track
      const track = r.name.toLowerCase();
      const trackData = pc.stressFilled[track];
      const trackSlots = getTrackTotalSlots(pc, track);
      if (trackData !== undefined && trackSlots) {
        const stressBar = document.createElement('div');
        stressBar.className = 'res-stress-mini-bar';
        const filled = trackData ? trackData.length : 0;
        const pct = Math.round((filled / trackSlots) * 100);
        stressBar.innerHTML = `<div class="res-stress-fill" style="width:${pct}%"></div>`;
        stressBar.title = `${filled}/${trackSlots} stress`;
        const stressNum = document.createElement('span');
        stressNum.className = 'res-stress-num';
        stressNum.textContent = filled + '/' + trackSlots;
        resRow.appendChild(nameSpan);
        resRow.appendChild(valBadge);
        resRow.appendChild(stressBar);
        resRow.appendChild(stressNum);
        resRow.appendChild(rollBtn);
      } else {
        resRow.appendChild(nameSpan);
        resRow.appendChild(valBadge);
        resRow.appendChild(rollBtn);
      }
      resBody.appendChild(resRow);
    });
    if (state.gmMode) {
      const addResBtn = document.createElement('button');
      addResBtn.textContent = '+ Add Resistance';
      addResBtn.addEventListener('click', () => {
        pc.resistances.push({ id: generateId('res'), name: '', value: 0 });
        appendLog('Added resistance', pc.id);
        saveAndRefresh();
      });
      resBody.appendChild(addResBtn);
    }

    // Bonds → left col
    const { sec: bondSec, body: bondBody } = makeSection('bonds', 'Bonds', 'bonds', pc);
    pc.bonds.forEach(bond => {
      const bRow = document.createElement('div');
      bRow.className = 'bond-row';
      const lvlSel = document.createElement('select');
      ['Individual','Street','Organisation'].forEach(lv => {
        const o = document.createElement('option');
        o.value = lv; o.textContent = lv;
        if (bond.level === lv) o.selected = true;
        lvlSel.appendChild(o);
      });
      lvlSel.addEventListener('change', e => {
        pushSectionUndo(pc, 'bonds', 'Edit bond level');
        bond.level = e.target.value;
        appendLog('Edited bond', pc.id);
        saveWithoutRefresh();
      });
      const nameIn = document.createElement('input');
      nameIn.type = 'text'; nameIn.placeholder = 'Bond name or person';
      nameIn.value = bond.name || '';
      nameIn.addEventListener('input', e => {
        bond.name = e.target.value;
        queueDeferredSave(`bond-name:${pc.id}:${bond.id}`);
      });
      nameIn.addEventListener('change', e => {
        pushSectionUndo(pc, 'bonds', 'Edit bond name');
        bond.name = e.target.value;
        appendLog('Edited bond', pc.id);
        saveWithoutRefresh();
      });
      const notesIn = document.createElement('input');
      notesIn.type = 'text'; notesIn.placeholder = 'Notes';
      notesIn.value = bond.notes || '';
      notesIn.addEventListener('input', e => {
        bond.notes = e.target.value;
        queueDeferredSave(`bond-notes:${pc.id}:${bond.id}`);
      });
      notesIn.addEventListener('change', e => {
        pushSectionUndo(pc, 'bonds', 'Edit bond notes');
        bond.notes = e.target.value;
        appendLog('Edited bond', pc.id);
        saveWithoutRefresh();
      });
      const remBtn = document.createElement('button');
      remBtn.textContent = '×'; remBtn.className = 'row-remove-btn';
      remBtn.addEventListener('click', () => {
        pushSectionUndo(pc, 'bonds', 'Remove bond');
        pc.bonds = pc.bonds.filter(b => b.id !== bond.id);
        saveAndRefresh();
      });
      bRow.appendChild(lvlSel); bRow.appendChild(nameIn); bRow.appendChild(notesIn); bRow.appendChild(remBtn);
      bondBody.appendChild(bRow);
    });
    if (!pc.bonds.length) {
      const emptyBond = document.createElement('p');
      emptyBond.className = 'text-muted';
      emptyBond.style.fontSize = '0.82rem';
      emptyBond.textContent = 'No bonds yet. Add allies, factions, or personal ties.';
      bondBody.appendChild(emptyBond);
      appendGettingStartedHelper(
        bondBody,
        `pc-bonds:${pc.id}`,
        'Add one Individual bond (person), one Street bond (community), and one Organisation bond to anchor roleplay choices.'
      );
    }
    const addBondBtn = document.createElement('button');
    addBondBtn.textContent = '+ Add Bond';
    addBondBtn.addEventListener('click', () => {
      pushSectionUndo(pc, 'bonds', 'Add bond');
      pc.bonds.push({ id: generateId('bond'), level: 'Individual', name: '', notes: '' });
      saveAndRefresh();
    });
    const undoBondBtn = document.createElement('button');
    undoBondBtn.type = 'button';
    undoBondBtn.textContent = 'Undo Bonds';
    undoBondBtn.disabled = !(pc.sectionUndo && pc.sectionUndo.bonds && pc.sectionUndo.bonds.length);
    undoBondBtn.title = sectionUndoTitle(pc, 'bonds', 'No bond changes to undo.');
    undoBondBtn.addEventListener('click', () => {
      undoSection(pc, 'bonds', 'No bond changes to undo.');
    });
    const redoBondBtn = document.createElement('button');
    redoBondBtn.type = 'button';
    redoBondBtn.textContent = 'Redo Bonds';
    redoBondBtn.disabled = !(pc.sectionRedo && pc.sectionRedo.bonds && pc.sectionRedo.bonds.length);
    redoBondBtn.title = sectionRedoTitle(pc, 'bonds', 'No bond changes to redo.');
    redoBondBtn.addEventListener('click', () => {
      redoSection(pc, 'bonds', 'No bond changes to redo.');
    });
    bondBody.appendChild(addBondBtn);
    bondBody.appendChild(undoBondBtn);
    bondBody.appendChild(redoBondBtn);
    leftCol.appendChild(bondSec);

    // Skills → right col (with Roll buttons and Mastered toggle)
    const { sec: skillsSec, body: skillsBody } = makeSection('skills', 'Skills', 'skills', pc);
    pc.skills.forEach(skill => {
      const row = createSkillRow(pc, skill, 'skills');
      skillsBody.appendChild(row);
    });
    const addSkillBtn = document.createElement('button');
    addSkillBtn.textContent = '+ Add Skill';
    addSkillBtn.addEventListener('click', async () => {
      const name = await askPrompt('Skill name:', '', { title: 'Add Skill', submitText: 'Add' });
      if (!name || !name.trim()) return;
      const trimmed = name.trim();
      if (pc.skills.some(s => s.name.toLowerCase() === trimmed.toLowerCase())) {
        showToast('Skill "' + trimmed + '" already exists.', 'warn'); return;
      }
      pc.skills.push({ id: generateId('skill'), name: trimmed, rating: 1, mastered: false });
      appendLog('Added skill', pc.id);
      saveAndRefresh();
    });
    skillsBody.appendChild(addSkillBtn);
    rightCol.appendChild(skillsSec);

    // Domains → right col
    const { sec: domainsSec, body: domainsBody } = makeSection('domains', 'Domains', 'domains', pc);
    pc.domains.forEach(d => {
      const row = createSkillRow(pc, d, 'domains');
      domainsBody.appendChild(row);
    });
    const addDomainBtn = document.createElement('button');
    addDomainBtn.textContent = '+ Add Domain';
    addDomainBtn.addEventListener('click', async () => {
      const name = await askPrompt('Domain name:', '', { title: 'Add Domain', submitText: 'Add' });
      if (!name || !name.trim()) return;
      const trimmed = name.trim();
      if (pc.domains.some(d => d.name.toLowerCase() === trimmed.toLowerCase())) {
        showToast('Domain "' + trimmed + '" already exists.', 'warn'); return;
      }
      pc.domains.push({ id: generateId('domain'), name: trimmed });
      appendLog('Added domain', pc.id);
      saveAndRefresh();
    });
    domainsBody.appendChild(addDomainBtn);
    rightCol.appendChild(domainsSec);
    rightCol.appendChild(resSec);

    const { sec: histSec, body: histBody } = makeSection('history', 'Action History', 'history', pc);
    const historyEntries = getRecentEntityActions(pc.id, 14);
    if (!historyEntries.length) {
      const none = document.createElement('p');
      none.className = 'text-muted';
      none.style.fontSize = '0.82rem';
      none.textContent = 'No recent actions recorded for this character.';
      histBody.appendChild(none);
      const openLogBtn = document.createElement('button');
      openLogBtn.className = 'toolbar-btn';
      openLogBtn.textContent = 'Open Activity Log';
      openLogBtn.addEventListener('click', () => {
        const logTab = document.querySelector('.tab-link[data-tab="log-view"]');
        if (logTab) logTab.click();
      });
      histBody.appendChild(openLogBtn);
    } else {
      const list = document.createElement('ul');
      list.className = 'entity-history-list';
      historyEntries.forEach((entry) => {
        const li = document.createElement('li');
        const time = document.createElement('time');
        time.dateTime = entry.time;
        const when = entry.time ? new Date(entry.time).toLocaleString() : 'Unknown time';
        const actor = entry.actor ? ` • ${entry.actor}` : '';
        time.textContent = when + actor;
        const text = document.createElement('span');
        text.textContent = ' ' + entry.action;
        li.appendChild(time);
        li.appendChild(text);
        list.appendChild(li);
      });
      histBody.appendChild(list);
      const openLogBtn = document.createElement('button');
      openLogBtn.className = 'toolbar-btn';
      openLogBtn.textContent = 'Open Full Log for This Character';
      openLogBtn.addEventListener('click', () => {
        logFilterState.query = entityLabel(pc);
        const input = document.getElementById('log-search-input');
        if (input) input.value = logFilterState.query;
        renderLog();
        const logTab = document.querySelector('.tab-link[data-tab="log-view"]');
        if (logTab) logTab.click();
      });
      histBody.appendChild(openLogBtn);
    }
    rightCol.appendChild(histSec);

    // Class features section: shows information and options derived from
    // the selected class. Includes refresh text, bond prompts, inventory
    // kit selection, core ability toggles and advances checkboxes.
    const classEff = CLASS_EFFECTS[pc.class];
    if (classEff) {
      const classTitle = pc.class ? `Class Features — ${pc.class}` : 'Class Features';
      const { sec: classSec, body: classBody } = makeSection('class-features', classTitle, null, pc);
      // Refresh condition with toggle
      const refreshDiv = document.createElement('div');
      refreshDiv.className = 'class-refresh' + (pc.refreshed ? ' refreshed' : '');
      const refreshLabel = document.createElement('span');
      refreshLabel.innerHTML = `<strong>Refresh:</strong> ${escapeHtml(classEff.refresh)}`;
      const refreshChk = document.createElement('button');
      refreshChk.className = 'refresh-toggle-btn' + (pc.refreshed ? ' active' : '');
      refreshChk.textContent = pc.refreshed ? '✓ Refreshed' : 'Mark Refreshed';
      refreshChk.title = 'Toggle whether core ability has been refreshed this scene';
      refreshChk.addEventListener('click', () => {
        pc.refreshed = !pc.refreshed;
        appendLog(pc.refreshed ? 'Marked as refreshed' : 'Cleared refresh', pc.id);
        saveAndRefresh();
      });
      const refreshApplyBtn = document.createElement('button');
      refreshApplyBtn.className = 'toolbar-btn';
      refreshApplyBtn.textContent = 'Apply Refresh (D3/D6/D8)';
      refreshApplyBtn.title = 'Roll D3/D6/D8 and clear stress per refresh rules';
      refreshApplyBtn.disabled = !canEditPC;
      refreshApplyBtn.addEventListener('click', () => openRefreshStressModal(pc));
      refreshDiv.appendChild(refreshLabel);
      refreshDiv.appendChild(refreshChk);
      refreshDiv.appendChild(refreshApplyBtn);
      classBody.appendChild(refreshDiv);
      // Bond prompts and responses
      if (classEff.bondPrompts && classEff.bondPrompts.length) {
        classEff.bondPrompts.forEach((prompt, idx) => {
          const promptDiv = document.createElement('div');
          promptDiv.className = 'class-bond-prompt';
          const label = document.createElement('label');
          label.textContent = prompt;
          label.style.display = 'block';
          const textarea = document.createElement('textarea');
          textarea.value = pc.classBondResponses[idx] || '';
          textarea.placeholder = 'Response...';
          textarea.addEventListener('input', e => {
            pc.classBondResponses[idx] = e.target.value;
            queueDeferredSave(`class-bond:${pc.id}:${idx}`);
          });
          textarea.addEventListener('change', e => {
            pc.classBondResponses[idx] = e.target.value;
            appendLog('Edited class bond response', pc.id);
            saveWithoutRefresh();
          });
          textarea.className = 'class-bond-response';
          promptDiv.appendChild(label);
          promptDiv.appendChild(textarea);
          classBody.appendChild(promptDiv);
        });
      }
      // Inventory kit selection — checkboxes so players can access all kits
      if (classEff.inventoryOptions && classEff.inventoryOptions.length) {
        const invGroup = document.createElement('div');
        invGroup.className = 'class-inv-options';
        const invLabel = document.createElement('div');
        invLabel.innerHTML = '<strong>Starting Equipment:</strong>';
        invLabel.style.marginBottom = '6px';
        invGroup.appendChild(invLabel);
        classEff.inventoryOptions.forEach(opt => {
          const optDiv = document.createElement('div');
          optDiv.className = 'class-inv-option';
          const wrap = document.createElement('label');
          wrap.style.display = 'flex';
          wrap.style.alignItems = 'flex-start';
          wrap.style.gap = '8px';
          wrap.style.width = '100%';
          const checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.style.marginTop = '3px';
          // An item kit is "selected" if ANY item from it exists with source classInv
          // We track selected kits as an array: pc.classInventorySelections (plural)
          if (!Array.isArray(pc.classInventorySelections)) {
            // Migrate from old single-selection model
            pc.classInventorySelections = pc.classInventorySelection ? [pc.classInventorySelection] : [];
          }
          checkbox.checked = pc.classInventorySelections.includes(opt.label);
          checkbox.addEventListener('change', e => {
            if (!Array.isArray(pc.classInventorySelections)) pc.classInventorySelections = [];
            if (e.target.checked) {
              if (!pc.classInventorySelections.includes(opt.label)) {
                pc.classInventorySelections.push(opt.label);
                // Add kit items without removing others
                const classEffNow = CLASS_EFFECTS[pc.class];
                if (classEffNow) {
                  const kitOpt = classEffNow.inventoryOptions.find(o => o.label === opt.label);
                  if (kitOpt) {
                    kitOpt.items.forEach(item => {
                      pc.inventory.push({
                        id: generateId('item'),
                        item: item.item,
                        quantity: item.quantity || 1,
                        type: item.type || 'other',
                        tags: item.tags ? item.tags.slice() : [],
                        notes: '',
                        stress: item.stress,
                        resistance: item.resistance,
                        source: `classInv:${pc.class}:${opt.label}`
                      });
                    });
                  }
                }
              }
            } else {
              pc.classInventorySelections = pc.classInventorySelections.filter(s => s !== opt.label);
              // Remove items from this specific kit
              pc.inventory = pc.inventory.filter(it => it.source !== `classInv:${pc.class}:${opt.label}`);
            }
            appendLog('Changed class kit', pc.id);
            saveAndRefresh();
          });
          const textContent = document.createElement('div');
          textContent.style.flex = '1';
          const optLabelEl = document.createElement('span');
          optLabelEl.textContent = opt.label;
          textContent.appendChild(optLabelEl);
          // List items in kit
          if (opt.items && opt.items.length) {
            const itemList = document.createElement('div');
            itemList.style.fontSize = '0.8rem';
            itemList.style.color = 'var(--spire-muted)';
            itemList.style.marginTop = '3px';
            opt.items.forEach(it => {
              const itSpan = document.createElement('div');
              itSpan.textContent = `• ${it.item}${it.stress ? ' (' + it.stress + ')' : ''}${it.resistance !== undefined ? ' [Res ' + it.resistance + ']' : ''}`;
              itemList.appendChild(itSpan);
            });
            textContent.appendChild(itemList);
          }
          if (opt.note) {
            const note = document.createElement('small');
            note.textContent = opt.note;
            note.style.display = 'block';
            note.style.color = 'var(--spire-muted)';
            note.style.marginTop = '3px';
            textContent.appendChild(note);
          }
          wrap.appendChild(checkbox);
          wrap.appendChild(textContent);
          optDiv.appendChild(wrap);
          invGroup.appendChild(optDiv);
        });
        classBody.appendChild(invGroup);
      }
      // Core abilities toggles
      if (classEff.coreAbilities && classEff.coreAbilities.length) {
        const coreDiv = document.createElement('div');
        coreDiv.className = 'class-core-abilities';
        const coreHeader = document.createElement('div');
        coreHeader.innerHTML = '<strong>Core Abilities:</strong>';
        coreDiv.appendChild(coreHeader);
        classEff.coreAbilities.forEach(ab => {
          const abDiv = document.createElement('div');
          abDiv.className = 'class-core-ability';
          const checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.checked = pc.coreAbilitiesState[ab] || false;
          checkbox.addEventListener('change', e => {
            pc.coreAbilitiesState[ab] = e.target.checked;
            appendLog('Toggled core ability', pc.id);
            saveAndRefresh();
          });
          const clabel = document.createElement('label');
          clabel.textContent = ab;
          abDiv.appendChild(checkbox);
          abDiv.appendChild(clabel);
          coreDiv.appendChild(abDiv);
        });
        classBody.appendChild(coreDiv);
      }
      // Advances checkboxes
      if (classEff.advances) {
        const advDiv = document.createElement('div');
        advDiv.className = 'class-advances';
        advDiv.innerHTML = '<strong>Advances:</strong>';
        // Helper to render a list
        function renderAdvList(level, list, unlocked = true, showLock = false) {
          if (!list || !list.length) return;
          const lvlDiv = document.createElement('div');
          lvlDiv.className = 'class-adv-level' + (unlocked ? '' : ' locked');
          const header = document.createElement('h4');
          const tierName = level.charAt(0).toUpperCase() + level.slice(1);
          const costMap = { low: 1, medium: 2, high: 3 };
          const prereq = level === 'medium' ? ' (need 2 Low first)' : level === 'high' ? ' (need 2 Medium first)' : '';
          header.textContent = tierName + ' Advances — ' + costMap[level] + ' AP' + ((!unlocked && prereq) ? prereq : '');
          if (!unlocked) header.title = 'Complete prerequisites to unlock this tier';
          lvlDiv.appendChild(header);
          list.forEach(adv => {
            const advRow = document.createElement('div');
            advRow.className = 'class-adv-option' + (!unlocked ? ' locked-adv' : '');
            const chk = document.createElement('input');
            chk.type = 'checkbox';
            chk.checked = pc.advances.includes(adv);
            chk.disabled = !unlocked && !pc.advances.includes(adv);
            chk.addEventListener('change', e => {
              if (!unlocked && e.target.checked) { e.target.checked = false; showToast('Complete prerequisites first.', 'warn'); return; }
              if (e.target.checked) { if (!pc.advances.includes(adv)) pc.advances.push(adv); }
              else { pc.advances = pc.advances.filter(a => a !== adv); }
              appendLog('Toggled advance', pc.id);
              saveAndRefresh();
            });
            const lbl = document.createElement('label');
            lbl.textContent = adv;
            if (!unlocked) lbl.style.opacity = '0.5';
            advRow.appendChild(chk);
            advRow.appendChild(lbl);
            lvlDiv.appendChild(advRow);
          });
          advDiv.appendChild(lvlDiv);
        }
        // Advance points tracker
        const apRow = document.createElement('div');
        apRow.className = 'advance-points-row';
        apRow.innerHTML = '<label>Advance Points (AP):</label>';
        const apInput = document.createElement('input');
        apInput.type = 'number';
        apInput.min = '0';
        apInput.value = pc.advancePoints || 0;
        apInput.style.width = '60px';
        apInput.addEventListener('change', e => {
          pc.advancePoints = Math.max(0, parseInt(e.target.value) || 0);
          saveAndRefresh();
        });
        apRow.appendChild(apInput);
        advDiv.insertBefore(apRow, advDiv.firstChild);

        // Count taken advances per tier for gate
        const lowTaken = (classEff.advances.low || []).filter(a => pc.advances.includes(a)).length;
        const medTaken = (classEff.advances.medium || []).filter(a => pc.advances.includes(a)).length;
        const medUnlocked = lowTaken >= 2;
        const highUnlocked = medTaken >= 2;
        const prog = document.createElement('div');
        prog.className = 'text-muted';
        prog.style.fontSize = '0.82rem';
        prog.style.marginBottom = '6px';
        const nextUnlock = !medUnlocked
          ? `Need ${Math.max(0, 2 - lowTaken)} more Low advance(s) to unlock Medium`
          : !highUnlocked
            ? `Need ${Math.max(0, 2 - medTaken)} more Medium advance(s) to unlock High`
            : 'All tiers unlocked';
        prog.textContent = `Progress: Low ${lowTaken}/2, Medium ${medTaken}/2. ${nextUnlock}.`;
        advDiv.appendChild(prog);

        renderAdvList('low', classEff.advances.low, true, true);
        renderAdvList('medium', classEff.advances.medium, medUnlocked, medUnlocked);
        renderAdvList('high', classEff.advances.high, highUnlocked, highUnlocked);
        classBody.appendChild(advDiv);
      }
      rightCol.appendChild(classSec);
    }
    // Inventory → below bonds (left col)
    const { sec: invSecEl, body: invBody } = makeSection('inventory', 'Inventory', 'inventory', pc);
    pc.inventory.forEach(item => invBody.appendChild(createInventoryRow(pc, item)));
    if (!pc.inventory.length) {
      const emptyInv = document.createElement('p');
      emptyInv.className = 'text-muted';
      emptyInv.style.fontSize = '0.82rem';
      emptyInv.textContent = 'No inventory yet. Add gear, weapons, and armor.';
      invBody.appendChild(emptyInv);
      appendGettingStartedHelper(
        invBody,
        `pc-inventory:${pc.id}`,
        'Start with your class kit, then add one personal item and one consumable to cover common scenes.'
      );
    }
    const addItemBtn = document.createElement('button');
    addItemBtn.textContent = '+ Add Item';
    addItemBtn.addEventListener('click', () => {
      pushSectionUndo(pc, 'inventory', 'Add inventory item');
      pc.inventory.push({ id: generateId('item'), type: 'other', item: '', quantity: 1, tags: [], notes: '', stress: undefined, resistance: undefined });
      appendLog('Added inventory item', pc.id);
      saveAndRefresh();
    });
    const undoInvBtn = document.createElement('button');
    undoInvBtn.type = 'button';
    undoInvBtn.textContent = 'Undo Inventory';
    undoInvBtn.disabled = !(pc.sectionUndo && pc.sectionUndo.inventory && pc.sectionUndo.inventory.length);
    undoInvBtn.title = sectionUndoTitle(pc, 'inventory', 'No inventory changes to undo.');
    undoInvBtn.addEventListener('click', () => {
      undoSection(pc, 'inventory', 'No inventory changes to undo.');
    });
    const redoInvBtn = document.createElement('button');
    redoInvBtn.type = 'button';
    redoInvBtn.textContent = 'Redo Inventory';
    redoInvBtn.disabled = !(pc.sectionRedo && pc.sectionRedo.inventory && pc.sectionRedo.inventory.length);
    redoInvBtn.title = sectionRedoTitle(pc, 'inventory', 'No inventory changes to redo.');
    redoInvBtn.addEventListener('click', () => {
      redoSection(pc, 'inventory', 'No inventory changes to redo.');
    });
    invBody.appendChild(addItemBtn);
    invBody.appendChild(undoInvBtn);
    invBody.appendChild(redoInvBtn);
    leftCol.appendChild(invSecEl);

    // Tasks → below inventory (left col)
    const { sec: taskSecEl, body: taskBody } = makeSection('tasks', 'Tasks / Quests', 'tasks', pc);
    if (!campPerm.taskViewByPc) campPerm.taskViewByPc = {};
    if (!campPerm.taskSortByPc) campPerm.taskSortByPc = {};
    const taskView = campPerm.taskViewByPc[pc.id] || 'all';
    const taskSort = campPerm.taskSortByPc[pc.id] || 'none';
    const taskControls = document.createElement('div');
    taskControls.className = 'task-controls-row';
    taskControls.style.display = 'flex';
    taskControls.style.gap = '6px';
    taskControls.style.marginBottom = '6px';
    const taskFilterSel = document.createElement('select');
    [
      { value: 'all', label: 'All Tasks' },
      { value: 'open', label: 'Open Only' },
      { value: 'done', label: 'Done Only' },
      { value: 'urgent', label: 'Urgent Only' }
    ].forEach((optData) => {
      const opt = document.createElement('option');
      opt.value = optData.value;
      opt.textContent = optData.label;
      if (taskView === optData.value) opt.selected = true;
      taskFilterSel.appendChild(opt);
    });
    taskFilterSel.addEventListener('change', (e) => {
      campPerm.taskViewByPc[pc.id] = e.target.value || 'all';
      saveAndRefresh();
    });
    taskControls.appendChild(taskFilterSel);
    const taskSortSel = document.createElement('select');
    [
      { value: 'none', label: 'Sort: Manual' },
      { value: 'priority', label: 'Sort: Priority' },
      { value: 'due', label: 'Sort: Due Date' },
      { value: 'status', label: 'Sort: Status' }
    ].forEach((optData) => {
      const opt = document.createElement('option');
      opt.value = optData.value;
      opt.textContent = optData.label;
      if (taskSort === optData.value) opt.selected = true;
      taskSortSel.appendChild(opt);
    });
    taskSortSel.addEventListener('change', (e) => {
      campPerm.taskSortByPc[pc.id] = e.target.value || 'none';
      saveAndRefresh();
    });
    taskControls.appendChild(taskSortSel);
    const clearDoneBtn = document.createElement('button');
    clearDoneBtn.type = 'button';
    clearDoneBtn.textContent = 'Clear Done';
    clearDoneBtn.disabled = !canEditPC || !pc.tasks.some((t) => t.status === 'Done');
    clearDoneBtn.addEventListener('click', () => {
      pushSectionUndo(pc, 'tasks', 'Clear done tasks');
      const before = pc.tasks.length;
      pc.tasks = pc.tasks.filter((t) => t.status !== 'Done');
      if (pc.tasks.length !== before) {
        appendLog('Cleared done tasks', pc.id);
        saveAndRefresh();
      }
    });
    taskControls.appendChild(clearDoneBtn);
    const undoTaskBtn = document.createElement('button');
    undoTaskBtn.type = 'button';
    undoTaskBtn.textContent = 'Undo Tasks';
    undoTaskBtn.disabled = !(pc.sectionUndo && pc.sectionUndo.tasks && pc.sectionUndo.tasks.length);
    undoTaskBtn.title = sectionUndoTitle(pc, 'tasks', 'No task changes to undo.');
    undoTaskBtn.addEventListener('click', () => {
      undoSection(pc, 'tasks', 'No task changes to undo.');
    });
    const redoTaskBtn = document.createElement('button');
    redoTaskBtn.type = 'button';
    redoTaskBtn.textContent = 'Redo Tasks';
    redoTaskBtn.disabled = !(pc.sectionRedo && pc.sectionRedo.tasks && pc.sectionRedo.tasks.length);
    redoTaskBtn.title = sectionRedoTitle(pc, 'tasks', 'No task changes to redo.');
    redoTaskBtn.addEventListener('click', () => {
      redoSection(pc, 'tasks', 'No task changes to redo.');
    });
    taskControls.appendChild(undoTaskBtn);
    taskControls.appendChild(redoTaskBtn);
    const taskHelpBtn = document.createElement('button');
    taskHelpBtn.type = 'button';
    taskHelpBtn.textContent = '?';
    taskHelpBtn.title = 'Task shortcuts and keyboard help';
    taskHelpBtn.addEventListener('click', () => openShortcutHelpModal());
    taskControls.appendChild(taskHelpBtn);
    taskBody.appendChild(taskControls);
    const filteredTasks = pc.tasks.filter((task) => {
      if (taskView === 'open') return task.status !== 'Done';
      if (taskView === 'done') return task.status === 'Done';
      if (taskView === 'urgent') return task.priority === 'Urgent';
      return true;
    });
    const sortedTasks = filteredTasks.slice();
    if (taskSort === 'priority') {
      const rank = { Urgent: 0, High: 1, Normal: 2, Low: 3 };
      sortedTasks.sort((a, b) => (rank[a.priority] ?? 99) - (rank[b.priority] ?? 99));
    } else if (taskSort === 'due') {
      sortedTasks.sort((a, b) => {
        const ad = a.dueDate || '';
        const bd = b.dueDate || '';
        if (!ad && !bd) return 0;
        if (!ad) return 1;
        if (!bd) return -1;
        return ad.localeCompare(bd);
      });
    } else if (taskSort === 'status') {
      const rank = { 'To Do': 0, Doing: 1, Done: 2 };
      sortedTasks.sort((a, b) => (rank[a.status] ?? 99) - (rank[b.status] ?? 99));
    }
    sortedTasks.forEach(task => taskBody.appendChild(createTaskRow(pc, task)));
    if (!pc.tasks.length) {
      const emptyTask = document.createElement('p');
      emptyTask.className = 'text-muted';
      emptyTask.style.fontSize = '0.82rem';
      emptyTask.textContent = 'No tasks yet. Track objectives, leads, and quests here.';
      taskBody.appendChild(emptyTask);
      appendGettingStartedHelper(
        taskBody,
        `pc-tasks:${pc.id}`,
        'Create one immediate objective and one longer-term consequence task so the next scene has clear momentum.'
      );
    } else if (!filteredTasks.length) {
      const emptyFilteredTask = document.createElement('p');
      emptyFilteredTask.className = 'text-muted';
      emptyFilteredTask.style.fontSize = '0.82rem';
      emptyFilteredTask.textContent = 'No tasks match this filter.';
      taskBody.appendChild(emptyFilteredTask);
      const clearTaskFilterBtn = document.createElement('button');
      clearTaskFilterBtn.className = 'toolbar-btn';
      clearTaskFilterBtn.textContent = 'Clear Task Filter';
      clearTaskFilterBtn.addEventListener('click', () => {
        campPerm.taskViewByPc[pc.id] = 'all';
        saveAndRefresh();
      });
      taskBody.appendChild(clearTaskFilterBtn);
    }
    const addTaskBtn = document.createElement('button');
    addTaskBtn.textContent = '+ Add Task';
    addTaskBtn.addEventListener('click', () => {
      pushSectionUndo(pc, 'tasks', 'Add task');
      pc.tasks.push({ id: generateId('task'), title: '', status: 'To Do', priority: 'Normal', dueDate: '', notes: '' });
      appendLog('Added task', pc.id);
      saveAndRefresh();
    });
    taskBody.appendChild(addTaskBtn);
    leftCol.appendChild(taskSecEl);

    // Notes (full-width below columns)
    const notesSec = document.createElement('div');
    notesSec.className = 'inspector-section';
    notesSec.innerHTML = '<h3>Notes</h3>';
    const notesField = document.createElement('textarea');
    notesField.value = pc.notes || '';
    notesField.placeholder = 'Notes...';
    notesField.addEventListener('input', e => {
      pc.notes = e.target.value;
      queueDeferredSave(`pc-notes:${pc.id}`);
    });
    notesField.addEventListener('change', e => {
      pc.notes = e.target.value;
      appendLog('Edited notes', pc.id);
      saveWithoutRefresh();
    });
    notesSec.appendChild(notesField);
    container.appendChild(notesSec);

    if (state.gmMode) {
      const gmNotesSec = document.createElement('div');
      gmNotesSec.className = 'inspector-section';
      gmNotesSec.innerHTML = '<h3>GM Notes</h3>';
      const gmField = document.createElement('textarea');
      gmField.value = pc.gmNotes || '';
      gmField.placeholder = 'Secret notes visible only to GM...';
      gmField.addEventListener('input', e => {
        pc.gmNotes = e.target.value;
        queueDeferredSave(`pc-gm-notes:${pc.id}`);
      });
      gmField.addEventListener('change', e => {
        pc.gmNotes = e.target.value;
        appendLog('Edited GM notes', pc.id);
        saveWithoutRefresh();
      });
      gmNotesSec.appendChild(gmField);
      container.appendChild(gmNotesSec);
    }
  }

  /**
   * Render a sheet for an NPC. Includes role, affiliation, a single
   * "Bond Stress" track, inventory, fallout and notes.
   */
  function renderNPCSheet(npc) {
    const container = document.getElementById('sheet-container');
    container.innerHTML = '';

    // Title
    const title = document.createElement('div');
    title.className = 'sheet-title';
    title.innerHTML = `<h2>${escapeHtml(entityLabel(npc))}</h2>`;
    const printBtn = document.createElement('button');
    printBtn.className = 'print-btn';
    printBtn.textContent = 'Print';
    printBtn.addEventListener('click', () => printEntitySheet(npc.id));
    if (state.gmMode) {
      const printPlayerBtn = document.createElement('button');
      printPlayerBtn.className = 'print-btn';
      printPlayerBtn.textContent = 'Print Player Copy';
      printPlayerBtn.title = 'One-click player-safe print';
      printPlayerBtn.addEventListener('click', () => printEntitySheet(npc.id, { forcePlayerSafe: true }));
      title.appendChild(printPlayerBtn);
    }
    if (state.gmMode) {
      const dupBtn = document.createElement('button');
      dupBtn.className = 'print-btn';
      dupBtn.textContent = 'Duplicate';
      dupBtn.addEventListener('click', () => {
        const camp = currentCampaign();
        const clone = JSON.parse(JSON.stringify(npc));
        clone.id = generateId('npc');
        clone.name = (clone.name || 'Copy') + ' (copy)';
        camp.entities[clone.id] = clone;
        appendLog('Duplicated NPC', clone.id);
        saveAndRefresh();
        selectEntity(clone.id);
      });
      title.appendChild(dupBtn);
      const delBtn = document.createElement('button');
      delBtn.className = 'print-btn print-btn-danger';
      delBtn.textContent = 'Delete';
      delBtn.addEventListener('click', async () => {
        await confirmAndDeleteEntity(npc.id);
      });
      title.appendChild(delBtn);
    }
    title.appendChild(printBtn);
    container.appendChild(title);

    // Pending approval banner — shown to GM with approve/reject; shown to player as read-only notice
    if (npc.pendingApproval) {
      const banner = document.createElement('div');
      banner.className = 'pending-approval-banner';
      const msg = document.createElement('span');
      if (state.gmMode) {
        msg.textContent = '⏳ This NPC was submitted by a player and awaits your approval.';
        const actions = document.createElement('div');
        actions.className = 'pending-actions';
        const approveBtn = document.createElement('button');
        approveBtn.className = 'pending-approve-btn';
        approveBtn.textContent = 'Approve';
        approveBtn.addEventListener('click', () => {
          npc.pendingApproval = false;
          appendLog('Approved NPC', npc.id);
          saveAndRefresh();
        });
        const rejectBtn = document.createElement('button');
        rejectBtn.className = 'pending-reject-btn';
        rejectBtn.textContent = 'Reject';
        rejectBtn.addEventListener('click', async () => {
          const ok = await askConfirm(`Reject and delete "${entityLabel(npc)}"?`, 'Reject NPC');
          if (!ok) return;
          if (deleteEntity(npc.id, { actionLabel: 'Rejected NPC' })) saveAndRefresh();
        });
        actions.appendChild(approveBtn);
        actions.appendChild(rejectBtn);
        banner.appendChild(msg);
        banner.appendChild(actions);
      } else {
        msg.textContent = '⏳ Submitted — awaiting GM approval.';
        banner.appendChild(msg);
      }
      container.appendChild(banner);
    }

    // Identity
    const identity = document.createElement('div');
    identity.className = 'inspector-section';
    identity.innerHTML = '<h3>Identity</h3>';
    identity.appendChild(createInputField(npc, 'name', 'Name'));
    identity.appendChild(createInputField(npc, 'role', 'Role / Archetype'));
    identity.appendChild(createPortraitField(npc, 'Portrait'));

    // Affiliation dropdown
    const affField = document.createElement('div');
    affField.className = 'inspector-field';
    const affLabel = document.createElement('label');
    affLabel.textContent = 'Affiliated Organisation';
    const affSelect = document.createElement('select');
    const noneOpt = document.createElement('option');
    noneOpt.value = '';
    noneOpt.textContent = 'None';
    affSelect.appendChild(noneOpt);
    Object.values(currentCampaign().entities).forEach(ent => {
      if (ent.type === 'org') {
        const opt = document.createElement('option');
        opt.value = ent.id;
        opt.textContent = entityLabel(ent);
        if (npc.affiliation === ent.id) opt.selected = true;
        affSelect.appendChild(opt);
      }
    });
    affSelect.addEventListener('change', e => {
      npc.affiliation = e.target.value;
      appendLog('Changed affiliation', npc.id);
      saveWithoutRefresh();
    });
    affField.appendChild(affLabel);
    affField.appendChild(affSelect);
    identity.appendChild(affField);

    // Threat level
    const threatField = document.createElement('div');
    threatField.className = 'inspector-field';
    threatField.innerHTML = '<label>Threat Level</label>';
    const threatSel = document.createElement('select');
    ['Minor','Significant','Major'].forEach(t => {
      const o = document.createElement('option');
      o.value = t; o.textContent = t;
      if ((npc.threatLevel || 'Minor') === t) o.selected = true;
      threatSel.appendChild(o);
    });
    threatSel.addEventListener('change', e => {
      npc.threatLevel = e.target.value;
      appendLog('Edited threat level', npc.id);
      saveWithoutRefresh();
    });
    threatField.appendChild(threatSel);
    identity.appendChild(threatField);

    // Disposition
    const dispField = document.createElement('div');
    dispField.className = 'inspector-field';
    dispField.innerHTML = '<label>Disposition toward PCs</label>';
    const dispSel = document.createElement('select');
    ['Friendly','Neutral','Wary','Hostile','Unknown'].forEach(d => {
      const o = document.createElement('option');
      o.value = d; o.textContent = d;
      if ((npc.disposition || 'Neutral') === d) o.selected = true;
      dispSel.appendChild(o);
    });
    dispSel.addEventListener('change', e => {
      npc.disposition = e.target.value;
      appendLog('Edited disposition', npc.id);
      saveWithoutRefresh();
    });
    dispField.appendChild(dispSel);
    identity.appendChild(dispField);

    // Wants / Fears / Leverage
    ['wants','fears','leverage'].forEach(field => {
      const f = document.createElement('div');
      f.className = 'inspector-field';
      f.innerHTML = '<label>' + field.charAt(0).toUpperCase() + field.slice(1) + '</label>';
      const inp = document.createElement('input');
      inp.type = 'text';
      inp.placeholder = field === 'wants' ? 'What do they want?' : field === 'fears' ? 'What do they fear?' : 'What leverage do they have?';
      inp.value = npc[field] || '';
      inp.addEventListener('input', e => {
        npc[field] = e.target.value;
        queueDeferredSave(`npc-${field}:${npc.id}`);
      });
      inp.addEventListener('change', e => {
        npc[field] = e.target.value;
        appendLog(`Edited ${field}`, npc.id);
        saveWithoutRefresh();
      });
      f.appendChild(inp);
      identity.appendChild(f);
    });
    if (state.gmMode) {
      identity.appendChild(createHideFromPlayersField(npc));
    }

    container.appendChild(identity);

    // Bond Stress — single track with +/- slots and individually toggling pips
    const stressSection = document.createElement('div');
    stressSection.className = 'inspector-section';
    stressSection.innerHTML = '<h3>Bond Stress</h3>';

    // Initialise bond stress data if missing
    if (!npc.bondStressSlots) npc.bondStressSlots = 10;
    if (!Array.isArray(npc.bondStressFilled)) npc.bondStressFilled = [];

    const stressRow = document.createElement('div');
    stressRow.className = 'stress-row stress-bond';

    const stressLabel = document.createElement('span');
    stressLabel.className = 'stress-label';
    stressLabel.textContent = 'Bond';
    stressRow.appendChild(stressLabel);

    const controls = document.createElement('div');
    controls.className = 'stress-controls';

    const minusBtn = document.createElement('button');
    minusBtn.textContent = '−';
    minusBtn.title = 'Remove stress slot';
    minusBtn.addEventListener('click', () => {
      if (npc.bondStressSlots <= 1) return;
      npc.bondStressSlots--;
      npc.bondStressFilled = npc.bondStressFilled.filter(i => i < npc.bondStressSlots);
      appendLog('Reduced bond stress slots', npc.id);
      saveAndRefresh();
    });

    const slotCount = document.createElement('span');
    slotCount.className = 'stress-slot-count';
    slotCount.textContent = npc.bondStressSlots;

    const plusBtn = document.createElement('button');
    plusBtn.textContent = '+';
    plusBtn.title = 'Add stress slot';
    plusBtn.addEventListener('click', () => {
      if (npc.bondStressSlots >= 20) return;
      npc.bondStressSlots++;
      appendLog('Increased bond stress slots', npc.id);
      saveAndRefresh();
    });

    controls.appendChild(minusBtn);
    controls.appendChild(slotCount);
    controls.appendChild(plusBtn);
    stressRow.appendChild(controls);

    const pipsContainer = document.createElement('div');
    pipsContainer.className = 'stress-pips';
    for (let i = 0; i < npc.bondStressSlots; i++) {
      const pip = document.createElement('div');
      pip.className = 'pip npc-bond-pip' + (npc.bondStressFilled.includes(i) ? ' active' : '');
      pip.title = `Bond stress ${i + 1}`;
      makeInteractivePip(pip, `Bond stress ${i + 1}`);
      const idx = i;
      pip.addEventListener('click', () => {
        const currentMax = npc.bondStressFilled.length > 0 ? Math.max(...npc.bondStressFilled) : -1;
        if (idx <= currentMax) {
          npc.bondStressFilled = Array.from({ length: idx }, (_, k) => k);
        } else {
          npc.bondStressFilled = Array.from({ length: idx + 1 }, (_, k) => k);
        }
        appendLog(`Set bond stress to ${npc.bondStressFilled.length}`, npc.id);
        saveAndRefresh();
      });
      pipsContainer.appendChild(pip);
    }
    stressRow.appendChild(pipsContainer);
    stressSection.appendChild(stressRow);
    container.appendChild(stressSection);

    // Fallout
    const falloutSection = document.createElement('div');
    falloutSection.className = 'inspector-section';
    falloutSection.innerHTML = '<h3>Fallout</h3>';
    npc.fallout.forEach(f => {
      falloutSection.appendChild(createFalloutRow(npc, f));
    });
    if (!npc.fallout.length) {
      const empty = document.createElement('p');
      empty.className = 'text-muted';
      empty.style.fontSize = '0.82rem';
      empty.textContent = 'No fallout recorded for this NPC.';
      falloutSection.appendChild(empty);
      appendGettingStartedHelper(
        falloutSection,
        `npc-fallout:${npc.id}`,
        'Use Minor fallout for short complications and Moderate/Severe when the NPC meaningfully changes the scene stakes.'
      );
      const quickAdd = document.createElement('button');
      quickAdd.className = 'toolbar-btn';
      quickAdd.textContent = 'Create First Fallout';
      quickAdd.addEventListener('click', () => {
        npc.fallout.push({
          id: generateId('fallout'),
          type: 'Blood',
          severity: 'Minor',
          name: '',
          description: '',
          timestamp: new Date().toISOString()
        });
        appendLog('Added fallout', npc.id);
        saveAndRefresh();
      });
      falloutSection.appendChild(quickAdd);
    }
    const addFalloutBtn = document.createElement('button');
    addFalloutBtn.className = 'section-add-btn';
    addFalloutBtn.innerHTML = '<span class="material-icons" style="font-size:14px">add</span> Add Fallout';
    addFalloutBtn.addEventListener('click', () => {
      npc.fallout.push({
        id: generateId('fallout'),
        type: 'Blood',
        severity: 'Minor',
        name: '',
        description: '',
        timestamp: new Date().toISOString()
      });
      appendLog('Added fallout', npc.id);
      saveAndRefresh();
    });
    falloutSection.appendChild(addFalloutBtn);
    container.appendChild(falloutSection);

    // Inventory — initialise if missing
    if (!Array.isArray(npc.inventory)) npc.inventory = [];
    const invSec = document.createElement('div');
    invSec.className = 'inspector-section';
    invSec.innerHTML = '<h3>Inventory</h3>';
    npc.inventory.forEach(item => {
      invSec.appendChild(createInventoryRow(npc, item));
    });
    if (!npc.inventory.length) {
      const empty = document.createElement('p');
      empty.className = 'text-muted';
      empty.style.fontSize = '0.82rem';
      empty.textContent = 'No inventory recorded for this NPC.';
      invSec.appendChild(empty);
      appendGettingStartedHelper(
        invSec,
        `npc-inventory:${npc.id}`,
        'Add one signature item and one leverage item so this NPC can drive scenes beyond dialogue.'
      );
      const quickAddItem = document.createElement('button');
      quickAddItem.className = 'toolbar-btn';
      quickAddItem.textContent = 'Add Starter Item';
      quickAddItem.addEventListener('click', () => {
        npc.inventory.push({
          id: generateId('item'),
          type: 'other',
          item: '',
          quantity: 1,
          tags: [],
          notes: '',
          stress: undefined,
          resistance: undefined
        });
        appendLog('Added inventory item', npc.id);
        saveAndRefresh();
      });
      invSec.appendChild(quickAddItem);
    }
    const addItemBtn = document.createElement('button');
    addItemBtn.className = 'section-add-btn';
    addItemBtn.innerHTML = '<span class="material-icons" style="font-size:14px">add</span> Add Item';
    addItemBtn.addEventListener('click', () => {
      npc.inventory.push({
        id: generateId('item'),
        type: 'other',
        item: '',
        quantity: 1,
        tags: [],
        notes: '',
        stress: undefined,
        resistance: undefined
      });
      appendLog('Added inventory item', npc.id);
      saveAndRefresh();
    });
    invSec.appendChild(addItemBtn);
    container.appendChild(invSec);

    // Notes
    const notesSec = document.createElement('div');
    notesSec.className = 'inspector-section';
    notesSec.innerHTML = '<h3>Notes</h3>';
    const notesField = document.createElement('textarea');
    notesField.value = npc.notes || '';
    notesField.placeholder = 'Notes...';
    notesField.addEventListener('input', e => {
      npc.notes = e.target.value;
      queueDeferredSave(`npc-notes:${npc.id}`);
    });
    notesField.addEventListener('change', e => {
      npc.notes = e.target.value;
      appendLog('Edited notes', npc.id);
      saveWithoutRefresh();
    });
    notesSec.appendChild(notesField);
    container.appendChild(notesSec);

    // GM Notes
    if (state.gmMode) {
      const gmSec = document.createElement('div');
      gmSec.className = 'inspector-section';
      gmSec.innerHTML = '<h3>GM Notes</h3>';
      const gmField = document.createElement('textarea');
      gmField.value = npc.gmNotes || '';
      gmField.placeholder = 'Secret notes visible only to GM...';
      gmField.addEventListener('input', e => {
        npc.gmNotes = e.target.value;
        queueDeferredSave(`npc-gm-notes:${npc.id}`);
      });
      gmField.addEventListener('change', e => {
        npc.gmNotes = e.target.value;
        appendLog('Edited GM notes', npc.id);
        saveWithoutRefresh();
      });
      gmSec.appendChild(gmField);
      container.appendChild(gmSec);
    }
  }

  /**
   * Render a sheet for an Organisation. Includes description, members,
   * relationships and notes.
   */
  function renderOrgSheet(org) {
    const container = document.getElementById('sheet-container');
    container.innerHTML = '';
    const title = document.createElement('div');
    title.className = 'sheet-title';
    title.innerHTML = `<h2>${escapeHtml(entityLabel(org))}</h2>`;
    const printBtn = document.createElement('button');
    printBtn.className = 'print-btn';
    printBtn.textContent = 'Print';
    printBtn.addEventListener('click', () => {
      printEntitySheet(org.id);
    });
    title.appendChild(printBtn);
    if (state.gmMode) {
      const printPlayerBtn = document.createElement('button');
      printPlayerBtn.className = 'print-btn';
      printPlayerBtn.textContent = 'Print Player Copy';
      printPlayerBtn.title = 'One-click player-safe print';
      printPlayerBtn.addEventListener('click', () => printEntitySheet(org.id, { forcePlayerSafe: true }));
      title.appendChild(printPlayerBtn);
    }
    container.appendChild(title);
    // Org duplicate button
    if (state.gmMode) {
      const dupBtn = document.createElement('button');
      dupBtn.className = 'print-btn';
      dupBtn.textContent = 'Duplicate';
      dupBtn.addEventListener('click', () => {
        const camp = currentCampaign();
        const clone = JSON.parse(JSON.stringify(org));
        clone.id = generateId('org');
        clone.name = (clone.name || 'Copy') + ' (copy)';
        camp.entities[clone.id] = clone;
        saveAndRefresh();
        selectEntity(clone.id);
      });
      title.appendChild(dupBtn);
      const delBtn = document.createElement('button');
      delBtn.className = 'print-btn print-btn-danger';
      delBtn.textContent = 'Delete';
      delBtn.addEventListener('click', async () => {
        await confirmAndDeleteEntity(org.id);
      });
      title.appendChild(delBtn);
    }
    // Name/identity section so organisation names are editable
    const identitySec = document.createElement('div');
    identitySec.className = 'inspector-section';
    identitySec.innerHTML = '<h3>Identity</h3>';
    identitySec.appendChild(createInputField(org, 'name', 'Name'));

    // Reach
    const reachField = document.createElement('div');
    reachField.className = 'inspector-field';
    reachField.innerHTML = '<label>Reach</label>';
    const reachSel = document.createElement('select');
    ['Local','District','City-wide','Empire-wide'].forEach(r => {
      const o = document.createElement('option');
      o.value = r; o.textContent = r;
      if ((org.reach || 'Local') === r) o.selected = true;
      reachSel.appendChild(o);
    });
    reachSel.addEventListener('change', e => {
      org.reach = e.target.value;
      appendLog('Edited reach', org.id);
      saveWithoutRefresh();
    });
    reachField.appendChild(reachSel);
    identitySec.appendChild(reachField);

    // Ministry relation
    const minRelField = document.createElement('div');
    minRelField.className = 'inspector-field';
    minRelField.innerHTML = '<label>Ministry Relationship</label>';
    const minRelSel = document.createElement('select');
    ['Unknown','Allied','Puppet','Infiltrated','Resistant','Targeted'].forEach(r => {
      const o = document.createElement('option');
      o.value = r; o.textContent = r;
      if ((org.ministryRelation || 'Unknown') === r) o.selected = true;
      minRelSel.appendChild(o);
    });
    minRelSel.addEventListener('change', e => {
      org.ministryRelation = e.target.value;
      appendLog('Edited ministry relation', org.id);
      saveWithoutRefresh();
    });
    minRelField.appendChild(minRelSel);
    identitySec.appendChild(minRelField);
    if (state.gmMode) {
      identitySec.appendChild(createHideFromPlayersField(org));
    }

    container.appendChild(identitySec);

    const descSec = document.createElement('div');
    descSec.className = 'inspector-section';
    descSec.innerHTML = '<h3>Description</h3>';
    const descField = document.createElement('textarea');
    descField.value = org.description || '';
    descField.placeholder = 'Description...';
    descField.addEventListener('input', e => {
      org.description = e.target.value;
      queueDeferredSave(`org-desc:${org.id}`);
    });
    descField.addEventListener('change', e => {
      org.description = e.target.value;
      appendLog('Edited description', org.id);
      saveWithoutRefresh();
    });
    descSec.appendChild(descField);
    container.appendChild(descSec);
    // Members list
    const membersSec = document.createElement('div');
    membersSec.className = 'inspector-section';
    membersSec.innerHTML = '<h3>Members</h3>';
    const memberList = document.createElement('ul');
    memberList.style.listStyle = 'none';
    memberList.style.padding = '0';
    org.members.forEach(memId => {
      const member = currentCampaign().entities[memId];
      if (!member || (!state.gmMode && member.gmOnly)) return;
      const li = document.createElement('li');
      li.style.display = 'flex';
      li.style.alignItems = 'center';
      li.style.justifyContent = 'space-between';
      li.style.gap = '8px';
      const memberName = document.createElement('span');
      memberName.textContent = entityLabel(member);
      li.appendChild(memberName);
      li.dataset.id = memId;
      memberName.addEventListener('click', () => selectEntity(memId));
      if (state.gmMode) {
        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'row-remove-btn';
        removeBtn.textContent = '×';
        removeBtn.title = 'Remove member';
        removeBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          org.members = org.members.filter((id) => id !== memId);
          appendLog('Removed member', org.id);
          saveAndRefresh();
        });
        li.appendChild(removeBtn);
      }
      memberList.appendChild(li);
    });
    const noMembers = !memberList.children.length;
    if (noMembers) {
      const emptyMembers = document.createElement('p');
      emptyMembers.className = 'text-muted';
      emptyMembers.style.fontSize = '0.82rem';
      emptyMembers.textContent = 'No members yet. Add PCs/NPCs below.';
      membersSec.appendChild(emptyMembers);
      appendGettingStartedHelper(
        membersSec,
        `org-members:${org.id}`,
        'Add at least one visible member and one hidden influencer to make organisation relationships actionable in play.'
      );
    }
    membersSec.appendChild(memberList);
    // Member add via dropdown
    const addMemRow = document.createElement('div');
    addMemRow.style.display = 'flex';
    addMemRow.style.gap = '6px';
    addMemRow.style.marginTop = '6px';
    const memSelect = document.createElement('select');
    const memNone = document.createElement('option');
    memNone.value = ''; memNone.textContent = 'Add member…';
    memSelect.appendChild(memNone);
    Object.values(currentCampaign().entities)
      .filter(e => e.id !== org.id && e.type !== 'org' && !org.members.includes(e.id))
      .forEach(e => {
        const o = document.createElement('option');
        o.value = e.id;
        o.textContent = entityLabel(e) + ' (' + e.type.toUpperCase() + ')';
        memSelect.appendChild(o);
      });
    const addMemBtn = document.createElement('button');
    addMemBtn.textContent = 'Add';
    addMemBtn.addEventListener('click', () => {
      const val = memSelect.value;
      if (!val) return;
      if (!org.members.includes(val)) {
        org.members.push(val);
        newRelationship(currentCampaign(), val, org.id, 'Member Of');
        appendLog('Added member', org.id);
        saveAndRefresh();
      }
    });
    addMemRow.appendChild(memSelect);
    addMemRow.appendChild(addMemBtn);
    membersSec.appendChild(addMemRow);
    if (noMembers) {
      const quickMemberBtn = document.createElement('button');
      quickMemberBtn.className = 'toolbar-btn';
      quickMemberBtn.style.marginTop = '6px';
      quickMemberBtn.textContent = 'Pick First Member';
      quickMemberBtn.addEventListener('click', () => {
        memSelect.focus();
        if (memSelect.options.length > 1 && !memSelect.value) {
          memSelect.selectedIndex = 1;
        }
      });
      membersSec.appendChild(quickMemberBtn);
    }
    container.appendChild(membersSec);
    // Notes
    const notesSec = document.createElement('div');
    notesSec.className = 'inspector-section';
    notesSec.innerHTML = '<h3>Notes</h3>';
    const notesField = document.createElement('textarea');
    notesField.value = org.notes || '';
    notesField.placeholder = 'Notes...';
    notesField.addEventListener('input', e => {
      org.notes = e.target.value;
      queueDeferredSave(`org-notes:${org.id}`);
    });
    notesField.addEventListener('change', e => {
      org.notes = e.target.value;
      appendLog('Edited notes', org.id);
      saveWithoutRefresh();
    });
    notesSec.appendChild(notesField);
    container.appendChild(notesSec);
    // GM notes
    if (state.gmMode) {
      const gmSec = document.createElement('div');
      gmSec.className = 'inspector-section';
      gmSec.innerHTML = '<h3>GM Notes</h3>';
      const gmField = document.createElement('textarea');
      gmField.value = org.gmNotes || '';
      gmField.placeholder = 'Secret notes visible only to GM...';
      gmField.addEventListener('input', e => {
        org.gmNotes = e.target.value;
        queueDeferredSave(`org-gm-notes:${org.id}`);
      });
      gmField.addEventListener('change', e => {
        org.gmNotes = e.target.value;
        appendLog('Edited GM notes', org.id);
        saveWithoutRefresh();
      });
      gmSec.appendChild(gmField);
      container.appendChild(gmSec);
    }
  }

  /**
   * Create a simple input field for a property on an entity. The value
   * updates the entity in place. When in player mode, input is disabled
   * unless editing PCs is allowed (only for PC fields). The label text
   * appears above the input. This helper reduces boilerplate for many
   * text inputs.
   */
  function createInputField(ent, prop, labelText) {
    const wrapper = document.createElement('div');
    wrapper.className = 'inspector-field';
    const label = document.createElement('label');
    label.textContent = labelText;
    wrapper.appendChild(label);
    const input = document.createElement('input');
    input.type = 'text';
    input.value = ent[prop] || '';
    // Disable editing in player mode unless this is the player's own PC
    const camp = currentCampaign();
    const isOwnPC = !state.gmMode && ent.type === 'pc' && camp.playerOwnedPcId === ent.id && camp.allowPlayerEditing;
    if (!state.gmMode && !isOwnPC) {
      input.disabled = true;
    }
    if (!state.gmMode && ent.pendingApproval) {
      input.disabled = true;
    }
    const nameProp = prop === 'name' || prop === 'firstName' || prop === 'lastName';
    if (!nameProp) {
      input.addEventListener('input', e => {
        ent[prop] = e.target.value;
        queueDeferredSave(`field:${ent.id}:${prop}`);
      });
    }
    input.addEventListener('change', e => {
      ent[prop] = e.target.value;
      appendLog(`Changed ${prop}`, ent.id);
      if (nameProp) saveAndRefresh();
      else saveWithoutRefresh();
    });
    wrapper.appendChild(input);
    return wrapper;
  }

  function applyNpcTemplate(npc, templateKey = 'none') {
    const tpl = getEffectiveNpcTemplates(currentCampaign())[templateKey];
    if (!npc || !tpl || templateKey === 'none') return;
    npc.role = tpl.role || npc.role || 'NPC';
    npc.threatLevel = tpl.threatLevel || npc.threatLevel || 'Minor';
    npc.disposition = tpl.disposition || npc.disposition || 'Neutral';
    npc.wants = tpl.wants || npc.wants || '';
    npc.fears = tpl.fears || npc.fears || '';
    npc.leverage = tpl.leverage || npc.leverage || '';
    if (!Array.isArray(npc.inventory)) npc.inventory = [];
    if (Array.isArray(tpl.inventory) && tpl.inventory.length) {
      npc.inventory = tpl.inventory.map((item) => ({
        id: generateId('item'),
        item: item.item || '',
        quantity: item.quantity || 1,
        type: item.type || 'other',
        stress: item.stress,
        resistance: item.resistance,
        tags: Array.isArray(item.tags) ? item.tags.slice() : [],
        notes: ''
      }));
    }
  }

  function createHideFromPlayersField(ent) {
    const wrapper = document.createElement('div');
    wrapper.className = 'inspector-field hide-from-players-field';
    const label = document.createElement('label');
    label.textContent = 'Hide from players';
    const toggleBtn = document.createElement('button');
    toggleBtn.type = 'button';
    toggleBtn.className = 'refresh-toggle-btn hide-player-toggle-btn' + (ent.gmOnly ? ' active' : '');
    const refreshToggleUi = () => {
      toggleBtn.classList.toggle('active', !!ent.gmOnly);
      toggleBtn.textContent = ent.gmOnly ? 'Hidden from players' : 'Visible to players';
      toggleBtn.title = ent.gmOnly
        ? 'This entity is hidden from player lists and web'
        : 'This entity is visible to players';
    };
    refreshToggleUi();
    toggleBtn.addEventListener('click', () => {
      ent.gmOnly = !ent.gmOnly;
      refreshToggleUi();
      appendLog(ent.gmOnly ? 'Hidden from players' : 'Visible to players', ent.id);
      saveAndRefresh();
    });
    const hint = document.createElement('small');
    hint.className = 'text-muted';
    hint.style.fontSize = '0.78rem';
    hint.textContent = 'Hidden entities are excluded from player lists and the player conspiracy web.';
    wrapper.appendChild(label);
    wrapper.appendChild(toggleBtn);
    wrapper.appendChild(hint);
    return wrapper;
  }

  /**
   * Render a portrait frame with local image upload. Images are stored as a
   * data URL in ent.image so they remain in exported campaign data.
   */
  function createPortraitField(ent, labelText = 'Portrait') {
    const wrapper = document.createElement('div');
    wrapper.className = 'inspector-field portrait-field';
    const label = document.createElement('label');
    label.textContent = labelText;
    wrapper.appendChild(label);

    const frame = document.createElement('div');
    frame.className = 'portrait-frame';
    if (ent.image) {
      const img = document.createElement('img');
      img.src = ent.image;
      img.alt = `${entityLabel(ent)} portrait`;
      frame.appendChild(img);
    } else {
      const empty = document.createElement('div');
      empty.className = 'portrait-placeholder';
      empty.textContent = 'Insert Portrait';
      frame.appendChild(empty);
    }
    wrapper.appendChild(frame);

    const controls = document.createElement('div');
    controls.className = 'portrait-controls';
    const uploadBtn = document.createElement('button');
    uploadBtn.type = 'button';
    uploadBtn.textContent = ent.image ? 'Replace Image' : 'Upload Image';
    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.textContent = 'Remove';
    removeBtn.className = 'row-remove-btn';
    removeBtn.style.display = ent.image ? '' : 'none';

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = 'image/*';
    fileInput.style.display = 'none';

    const camp = currentCampaign();
    const isOwnPC = !state.gmMode && ent.type === 'pc' && camp.playerOwnedPcId === ent.id && camp.allowPlayerEditing;
    const readOnly = (!state.gmMode && !isOwnPC) || (!!ent.pendingApproval && !state.gmMode);
    if (readOnly) {
      uploadBtn.disabled = true;
      removeBtn.disabled = true;
    }

    uploadBtn.addEventListener('click', () => {
      if (readOnly) return;
      fileInput.click();
    });
    fileInput.addEventListener('change', (e) => {
      const file = e.target.files && e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (evt) => {
        ent.image = evt.target.result;
        appendLog('Updated portrait', ent.id);
        saveAndRefresh();
      };
      reader.readAsDataURL(file);
      e.target.value = '';
    });
    removeBtn.addEventListener('click', () => {
      if (readOnly || !ent.image) return;
      ent.image = '';
      appendLog('Removed portrait', ent.id);
      saveAndRefresh();
    });

    controls.appendChild(uploadBtn);
    controls.appendChild(removeBtn);
    controls.appendChild(fileInput);
    wrapper.appendChild(controls);
    return wrapper;
  }

  function totalArmorResistance(pc) {
    if (!pc || !Array.isArray(pc.inventory)) return 0;
    return pc.inventory.reduce((sum, item) => {
      if (item.type !== 'armor') return sum;
      return sum + (parseInt(item.resistance, 10) || 0);
    }, 0);
  }

  function getTrackBaseSlots(pc, track) {
    if (!pc || !pc.stressSlots) return 10;
    return Math.max(1, parseInt(pc.stressSlots[track], 10) || 10);
  }

  function getTrackFreeSlots(pc, track) {
    if (!pc || !track) return 0;
    const key = String(track).toLowerCase();
    let free = 0;
    if (Array.isArray(pc.resistances)) {
      pc.resistances.forEach((r) => {
        if (!r || !r.name) return;
        if (String(r.name).toLowerCase() !== key) return;
        free += Math.max(0, parseInt(r.value, 10) || 0);
      });
    }
    if (key === 'blood') free += totalArmorResistance(pc);
    return Math.max(0, free);
  }

  function getTrackTotalSlots(pc, track) {
    return getTrackBaseSlots(pc, track) + getTrackFreeSlots(pc, track);
  }

  function getTrackFalloutCountedStress(pc, track) {
    const filled = (pc && pc.stressFilled && pc.stressFilled[track]) ? pc.stressFilled[track].length : 0;
    const free = getTrackFreeSlots(pc, track);
    return Math.min(10, Math.max(0, filled - free));
  }

  function openRefreshStressModal(pc) {
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    const content = document.getElementById('modal-content');
    const titleEl = document.getElementById('modal-title');
    if (titleEl) titleEl.textContent = 'Apply Refresh Stress Removal';
    content.innerHTML = '';

    const form = document.createElement('div');
    form.className = 'modal-form';

    const intro = document.createElement('p');
    intro.className = 'text-muted';
    intro.style.margin = '0 0 10px';
    intro.textContent = 'Roll D3, D6, or D8 based on how fully the refresh condition was met.';
    form.appendChild(intro);

    const trackField = document.createElement('div');
    trackField.className = 'modal-field';
    const trackLabel = document.createElement('label');
    trackLabel.textContent = 'Clear From';
    const trackSel = document.createElement('select');
    const autoOpt = document.createElement('option');
    autoOpt.value = '__auto__';
    autoOpt.textContent = 'Highest Stress Track';
    trackSel.appendChild(autoOpt);
    ['blood', 'mind', 'silver', 'shadow', 'reputation'].forEach((t) => {
      const opt = document.createElement('option');
      opt.value = t;
      opt.textContent = t.charAt(0).toUpperCase() + t.slice(1);
      trackSel.appendChild(opt);
    });
    trackField.appendChild(trackLabel);
    trackField.appendChild(trackSel);
    form.appendChild(trackField);

    const actions = document.createElement('div');
    actions.style.display = 'flex';
    actions.style.gap = '8px';
    actions.style.flexWrap = 'wrap';
    actions.style.marginTop = '10px';

    function resolveTrack() {
      if (trackSel.value !== '__auto__') return trackSel.value;
      const tracks = ['blood', 'mind', 'silver', 'shadow', 'reputation'];
      let best = null;
      let bestVal = -1;
      tracks.forEach((t) => {
        const v = (pc.stressFilled && pc.stressFilled[t]) ? pc.stressFilled[t].length : 0;
        if (v > bestVal) {
          bestVal = v;
          best = t;
        }
      });
      return best;
    }

    function applyRefreshRoll(sides) {
      const targetTrack = resolveTrack();
      if (!targetTrack) {
        showToast('No stress tracks available to clear.', 'warn');
        return;
      }
      const current = (pc.stressFilled && pc.stressFilled[targetTrack]) ? pc.stressFilled[targetTrack].length : 0;
      if (current <= 0) {
        showToast('Selected track has no stress to clear.', 'warn');
        return;
      }
      const rolled = Math.ceil(Math.random() * sides);
      const next = Math.max(0, current - rolled);
      const removed = current - next;
      setPCStressLevel(pc, targetTrack, next, { triggerFallout: false });
      pc.refreshed = true;
      appendLog(`Applied refresh: cleared ${removed} ${targetTrack} stress (D${sides}=${rolled})`, pc.id);
      closeModal();
      saveAndRefresh();
    }

    [
      { sides: 3, label: 'Roll D3' },
      { sides: 6, label: 'Roll D6' },
      { sides: 8, label: 'Roll D8' }
    ].forEach((die) => {
      const btn = document.createElement('button');
      btn.className = 'modal-submit';
      btn.textContent = die.label;
      btn.addEventListener('click', () => applyRefreshRoll(die.sides));
      actions.appendChild(btn);
    });

    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    cancelBtn.addEventListener('click', () => closeModal());
    actions.appendChild(cancelBtn);

    form.appendChild(actions);
    content.appendChild(form);
    overlay.classList.remove('hidden');
    modal.classList.remove('hidden');
  }

  function openStressApplicationModal(pc) {
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    const content = document.getElementById('modal-content');
    const titleEl = document.getElementById('modal-title');
    if (titleEl) titleEl.textContent = 'Apply Stress';
    content.innerHTML = '';

    const form = document.createElement('div');
    form.className = 'modal-form';

    const sourceField = document.createElement('div');
    sourceField.className = 'modal-field';
    const sourceLabel = document.createElement('label');
    sourceLabel.textContent = 'Source';
    const sourceSel = document.createElement('select');
    ['Weapon', 'Fallout', 'Consequence', 'Cost', 'Other'].forEach((s) => {
      const opt = document.createElement('option');
      opt.value = s;
      opt.textContent = s;
      sourceSel.appendChild(opt);
    });
    sourceField.appendChild(sourceLabel);
    sourceField.appendChild(sourceSel);
    form.appendChild(sourceField);

    const trackField = document.createElement('div');
    trackField.className = 'modal-field';
    const trackLabel = document.createElement('label');
    trackLabel.textContent = 'Resistance Track';
    const trackSel = document.createElement('select');
    ['blood','mind','silver','shadow','reputation'].forEach((t) => {
      const opt = document.createElement('option');
      opt.value = t;
      opt.textContent = t.charAt(0).toUpperCase() + t.slice(1);
      trackSel.appendChild(opt);
    });
    trackField.appendChild(trackLabel);
    trackField.appendChild(trackSel);
    form.appendChild(trackField);

    const amountField = document.createElement('div');
    amountField.className = 'modal-field';
    const amountLabel = document.createElement('label');
    amountLabel.textContent = 'Incoming Stress';
    const amountInput = document.createElement('input');
    amountInput.type = 'number';
    amountInput.min = '0';
    amountInput.value = '1';
    amountField.appendChild(amountLabel);
    amountField.appendChild(amountInput);
    form.appendChild(amountField);

    const armorHint = document.createElement('div');
    armorHint.className = 'text-muted';
    armorHint.style.fontSize = '0.82rem';
    armorHint.style.marginTop = '-4px';
    armorHint.textContent = `Armor grants ${totalArmorResistance(pc)} free Blood slot(s); it does not directly reduce incoming stress.`;
    form.appendChild(armorHint);

    const preview = document.createElement('div');
    preview.className = 'text-muted';
    preview.style.fontSize = '0.84rem';
    preview.style.marginTop = '2px';
    form.appendChild(preview);

    function refreshPreview() {
      const incoming = Math.max(0, parseInt(amountInput.value, 10) || 0);
      preview.textContent = `Applied stress: ${incoming}`;
    }
    [amountInput].forEach((el) => {
      el.addEventListener('input', refreshPreview);
      el.addEventListener('change', refreshPreview);
    });
    refreshPreview();

    const btnRow = document.createElement('div');
    btnRow.style.display = 'flex';
    btnRow.style.gap = '8px';
    btnRow.style.marginTop = '10px';
    const applyBtn = document.createElement('button');
    applyBtn.className = 'modal-submit';
    applyBtn.textContent = 'Apply';
    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    btnRow.appendChild(applyBtn);
    btnRow.appendChild(cancelBtn);
    form.appendChild(btnRow);

    applyBtn.addEventListener('click', () => {
      const track = trackSel.value;
      const incoming = Math.max(0, parseInt(amountInput.value, 10) || 0);
      const current = (pc.stressFilled && pc.stressFilled[track]) ? pc.stressFilled[track].length : 0;
      setPCStressLevel(pc, track, current + incoming);
      appendLog(`Applied ${incoming} ${track} stress from ${sourceSel.value}`, pc.id);
      closeModal();
      saveAndRefresh();
    });
    cancelBtn.addEventListener('click', () => closeModal());

    content.appendChild(form);
    overlay.classList.remove('hidden');
    modal.classList.remove('hidden');
  }

  function openFalloutLookupModal(track = 'Blood', severity = 'Minor') {
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    const content = document.getElementById('modal-content');
    const titleEl = document.getElementById('modal-title');
    if (titleEl) titleEl.textContent = `Fallout Lookup: ${track} (${severity})`;
    content.innerHTML = '';

    const wrap = document.createElement('div');
    wrap.className = 'fallout-lookup-wrap';
    const intro = document.createElement('p');
    intro.className = 'text-muted';
    intro.textContent = 'Use these as fast in-session prompts. Rename/edit to fit fiction.';
    wrap.appendChild(intro);

    const camp = currentCampaign();
    const guidance = (camp && camp.falloutGuidance) ? camp.falloutGuidance : defaultFalloutGuidance();
    const table = guidance[track] || guidance.Blood || defaultFalloutGuidance().Blood;
    ['Minor', 'Moderate', 'Severe'].forEach((level) => {
      const block = document.createElement('div');
      block.className = 'fallout-lookup-block' + (level === severity ? ' active' : '');
      const header = document.createElement('strong');
      header.textContent = level;
      block.appendChild(header);
      const ul = document.createElement('ul');
      ul.className = 'fallout-lookup-list';
      (table[level] || []).forEach((item) => {
        const li = document.createElement('li');
        li.textContent = item;
        ul.appendChild(li);
      });
      block.appendChild(ul);
      wrap.appendChild(block);
    });

    content.appendChild(wrap);
    overlay.classList.remove('hidden');
    modal.classList.remove('hidden');
  }

  /**
   * Helper to create a fallout row for PCs and NPCs. Allows editing of
   * type, severity, name and description inline. Remove button deletes
   * the fallout entry.
   */
  function createFalloutRow(ent, f) {
    const row = document.createElement('div');
    row.className = 'fallout-row';
    row.style.marginBottom = '4px';
    // Type select
    const typeSel = document.createElement('select');
    ['Blood','Mind','Silver','Shadow','Reputation'].forEach(t => {
      const opt = document.createElement('option');
      opt.value = t;
      opt.textContent = t;
      if (f.type === t) opt.selected = true;
      typeSel.appendChild(opt);
    });
    typeSel.addEventListener('change', e => {
      f.type = e.target.value;
      appendLog('Edited fallout', ent.id);
      saveAndRefresh();
    });
    row.appendChild(typeSel);
    // Severity
    const sevSel = document.createElement('select');
    ['Minor','Moderate','Severe'].forEach(s => {
      const opt = document.createElement('option');
      opt.value = s;
      opt.textContent = s;
      if (f.severity === s) opt.selected = true;
      sevSel.appendChild(opt);
    });
    sevSel.addEventListener('change', e => {
      f.severity = e.target.value;
      appendLog('Edited fallout', ent.id);
      saveAndRefresh();
    });
    row.appendChild(sevSel);
    // Name
    const nameInput = document.createElement('input');
    nameInput.type = 'text';
    nameInput.placeholder = 'Name';
    nameInput.value = f.name || '';
    nameInput.style.flex = '1';
    nameInput.addEventListener('change', e => {
      f.name = e.target.value;
      appendLog('Edited fallout', ent.id);
      saveAndRefresh();
    });
    row.appendChild(nameInput);
    // Description
    const descInput = document.createElement('input');
    descInput.type = 'text';
    descInput.placeholder = 'Description';
    descInput.value = f.description || '';
    descInput.style.flex = '2';
    descInput.addEventListener('change', e => {
      f.description = e.target.value;
      appendLog('Edited fallout', ent.id);
      saveAndRefresh();
    });
    row.appendChild(descInput);
    // Quick fallout lookup
    const lookupBtn = document.createElement('button');
    lookupBtn.textContent = 'Lookup';
    lookupBtn.className = 'fallout-lookup-btn';
    lookupBtn.title = 'Open fallout guidance table';
    lookupBtn.addEventListener('click', () => {
      openFalloutLookupModal(f.type || 'Blood', f.severity || 'Minor');
    });
    row.appendChild(lookupBtn);
    // Resolved toggle
    const resolveBtn = document.createElement('button');
    resolveBtn.className = 'fallout-resolve-btn' + (f.resolved ? ' resolved' : '');
    resolveBtn.textContent = f.resolved ? '✓ Resolved' : 'Resolve';
    resolveBtn.title = f.resolved ? 'Mark as active again' : 'Mark as resolved';
    resolveBtn.addEventListener('click', async () => {
      f.resolved = !f.resolved;
      appendLog(f.resolved ? 'Resolved fallout' : 'Reopened fallout', ent.id);
      if (f.resolved && ent.type === 'pc') {
        const track = String(f.type || '').toLowerCase();
        const stressTracks = ['blood', 'mind', 'silver', 'shadow', 'reputation'];
        if (stressTracks.includes(track) && ent.stressFilled && Array.isArray(ent.stressFilled[track])) {
          const clearAmount = stressClearAmountForSeverity(f.severity || 'Minor');
          const current = ent.stressFilled[track].length;
          if (current > 0) {
            const suggested = Math.max(0, current - clearAmount);
            const ok = await askConfirm(
              `Clear ${clearAmount} ${track} stress for resolved ${f.severity || 'Minor'} fallout? (${current} → ${suggested})`,
              'Resolve Fallout'
            );
            if (ok) {
              setPCStressLevel(ent, track, suggested, { triggerFallout: false });
              appendLog('Cleared stress from resolved fallout', ent.id);
            }
          }
        }
      }
      saveAndRefresh();
    });
    row.appendChild(resolveBtn);
    // Remove button
    const remBtn = document.createElement('button');
    remBtn.textContent = '×';
    remBtn.className = 'row-remove-btn';
    remBtn.addEventListener('click', () => {
      ent.fallout = ent.fallout.filter(x => x.id !== f.id);
      appendLog('Removed fallout', ent.id);
      saveAndRefresh();
    });
    row.appendChild(remBtn);
    return row;
  }

  /**
   * Helper to create a skill/domain/resistance row. Accepts an object
   * with id, name and rating/value. The listName argument indicates
   * which property array on the entity this row belongs to.
   */
  function createSkillRow(ent, obj, listName) {
    const row = document.createElement('div');
    row.className = 'skill-row';
    row.style.display = 'flex';
    row.style.alignItems = 'center';
    row.style.marginBottom = '4px';

    // Utility: remove button for list rows
    const createRemoveButton = () => {
      const remBtn = document.createElement('button');
      remBtn.textContent = '×';
      remBtn.className = 'row-remove-btn';
      remBtn.addEventListener('click', () => {
        ent[listName] = ent[listName].filter(x => x.id !== obj.id);
        appendLog('Removed item', ent.id);
        saveAndRefresh();
      });
      return remBtn;
    };

    if (listName === 'skills') {
      const nameSelect = document.createElement('select');
      nameSelect.style.flex = '2';
      SKILL_OPTIONS.forEach(opt => {
        const option = document.createElement('option');
        option.value = opt; option.textContent = opt;
        if (obj.name === opt) option.selected = true;
        nameSelect.appendChild(option);
      });
      nameSelect.addEventListener('change', e => {
        obj.name = e.target.value;
        appendLog('Edited skill', ent.id);
        saveAndRefresh();
      });
      // Mastered toggle (replaces numeric rating in Spire)
      const masteredBtn = document.createElement('button');
      masteredBtn.className = 'mastered-toggle' + (obj.mastered ? ' active' : '');
      masteredBtn.textContent = obj.mastered ? '★ Mastered' : '☆ Skilled';
      masteredBtn.title = 'Toggle Mastered (Mastered = extra die on rolls)';
      masteredBtn.addEventListener('click', () => {
        obj.mastered = !obj.mastered;
        obj.rating = obj.mastered ? 2 : 1;
        appendLog('Toggled skill mastered', ent.id);
        saveAndRefresh();
      });
      // Roll button
      const rollBtn = document.createElement('button');
      rollBtn.className = 'skill-roll-btn';
      rollBtn.title = 'Roll ' + (obj.name || 'skill');
      rollBtn.innerHTML = '⚄';
      rollBtn.addEventListener('click', () => {
        openDiceRollerForSkill(obj.name, obj.mastered, ent.type === 'pc' ? ent : null);
      });
      row.appendChild(nameSelect);
      row.appendChild(masteredBtn);
      row.appendChild(rollBtn);
      row.appendChild(createRemoveButton());
    } else if (listName === 'domains') {
      // Domains are binary in Spire: relevant or not.
      const nameSelect = document.createElement('select');
      nameSelect.style.flex = '2';
      DOMAIN_OPTIONS.forEach(opt => {
        const option = document.createElement('option');
        option.value = opt;
        option.textContent = opt;
        if (obj.name === opt) option.selected = true;
        nameSelect.appendChild(option);
      });
      nameSelect.addEventListener('change', e => {
        obj.name = e.target.value;
        appendLog('Edited domain', ent.id);
        saveAndRefresh();
      });
      row.appendChild(nameSelect);
      row.appendChild(createRemoveButton());
    } else {
      // Default behaviour for resistances and other generic lists: free text name and numeric value.
      const nameInput = document.createElement('input');
      nameInput.type = 'text';
      nameInput.placeholder = 'Name';
      nameInput.value = obj.name || '';
      nameInput.style.flex = '2';
      nameInput.addEventListener('change', e => {
        obj.name = e.target.value;
        appendLog('Edited attribute', ent.id);
        saveAndRefresh();
      });
      const valInput = document.createElement('input');
      valInput.type = 'number';
      valInput.value = obj.rating !== undefined ? obj.rating : (obj.value || 0);
      valInput.style.width = '50px';
      valInput.addEventListener('change', e => {
        const val = parseInt(e.target.value) || 0;
        if (obj.rating !== undefined) obj.rating = val; else obj.value = val;
        appendLog('Edited attribute value', ent.id);
        saveAndRefresh();
      });
      row.appendChild(nameInput);
      row.appendChild(valInput);
      row.appendChild(createRemoveButton());
    }
    return row;
  }

  /**
   * Helper to create an inventory row. Handles editing item name, quantity,
   * tags and notes.
   */
  function createInventoryRow(pc, item) {
    const row = document.createElement('div');
    row.className = 'inventory-row';
    row.style.display = 'flex';
    row.style.flexWrap = 'wrap';
    row.style.alignItems = 'center';
    row.style.marginBottom = '4px';

    // Ensure the item has type and other fields set
    if (!item.type) item.type = 'other';
    if (!Array.isArray(item.tags)) item.tags = [];

    // Dropdown for item type
    const typeSelect = document.createElement('select');
    ['other', 'weapon', 'armor'].forEach(t => {
      const opt = document.createElement('option');
      opt.value = t;
      opt.textContent = t.charAt(0).toUpperCase() + t.slice(1);
      if (item.type === t) opt.selected = true;
      typeSelect.appendChild(opt);
    });
    typeSelect.style.flex = '1';
    typeSelect.addEventListener('change', e => {
      pushSectionUndo(pc, 'inventory', 'Change inventory type');
      item.type = e.target.value;
      // Reset type-specific fields
      if (item.type === 'weapon') {
        item.stress = item.stress || WEAPON_STRESS_LEVELS[0];
        item.resistance = undefined;
        item.tags = item.tags.filter(tag => WEAPON_TAGS.includes(tag));
      } else if (item.type === 'armor') {
        item.resistance = item.resistance || 0;
        item.stress = undefined;
        item.tags = item.tags.filter(tag => ARMOR_TAGS.includes(tag));
      } else {
        item.stress = undefined;
        item.resistance = undefined;
      }
      appendLog('Changed item type', pc.id);
      saveAndRefresh();
    });

    // Basic item name input
    const itemInput = document.createElement('input');
    itemInput.type = 'text';
    itemInput.placeholder = 'Item';
    itemInput.value = item.item || '';
    itemInput.style.flex = '2';
    itemInput.addEventListener('input', e => {
      item.item = e.target.value;
      queueDeferredSave(`inv-item:${pc.id}:${item.id}`);
    });
    itemInput.addEventListener('change', e => {
      pushSectionUndo(pc, 'inventory', 'Edit inventory item');
      item.item = e.target.value;
      appendLog('Edited inventory', pc.id);
      saveWithoutRefresh();
    });

    // Quantity input
    const qtyInput = document.createElement('input');
    qtyInput.type = 'number';
    qtyInput.value = item.quantity || 1;
    qtyInput.min = '1';
    qtyInput.style.width = '60px';
    qtyInput.addEventListener('change', e => {
      pushSectionUndo(pc, 'inventory', 'Edit inventory quantity');
      item.quantity = parseInt(e.target.value) || 1;
      appendLog('Edited inventory', pc.id);
      saveWithoutRefresh();
    });

    // Notes input
    const notesInput = document.createElement('input');
    notesInput.type = 'text';
    notesInput.placeholder = 'Notes';
    notesInput.value = item.notes || '';
    notesInput.style.flex = '3';
    notesInput.addEventListener('input', e => {
      item.notes = e.target.value;
      queueDeferredSave(`inv-notes:${pc.id}:${item.id}`);
    });
    notesInput.addEventListener('change', e => {
      pushSectionUndo(pc, 'inventory', 'Edit inventory notes');
      item.notes = e.target.value;
      appendLog('Edited inventory', pc.id);
      saveWithoutRefresh();
    });

    // Container for type-specific fields (stress/resistance and tags)
    const typeFields = document.createElement('div');
    typeFields.style.display = 'flex';
    typeFields.style.flexWrap = 'wrap';
    typeFields.style.alignItems = 'center';
    typeFields.style.gap = '4px';

    function renderTypeFields() {
      // Clear existing fields
      typeFields.innerHTML = '';
      if (item.type === 'weapon') {
        // Stress level select
        const stressLabel = document.createElement('label');
        stressLabel.textContent = 'Stress:';
        const stressSelect = document.createElement('select');
        WEAPON_STRESS_LEVELS.forEach(s => {
          const opt = document.createElement('option');
          opt.value = s;
          opt.textContent = s;
          if (item.stress === s) opt.selected = true;
          stressSelect.appendChild(opt);
        });
        stressSelect.addEventListener('change', e => {
          pushSectionUndo(pc, 'inventory', 'Edit weapon stress');
          item.stress = e.target.value;
          appendLog('Edited weapon stress', pc.id);
          saveAndRefresh();
        });
        // Weapon tags — toggleable pills (multiple allowed, repeatable)
        const tagsWrapper = document.createElement('div');
        tagsWrapper.style.display = 'flex';
        tagsWrapper.style.flexWrap = 'wrap';
        tagsWrapper.style.gap = '4px';
        tagsWrapper.style.marginTop = '6px';
        const tagsHeader = document.createElement('div');
        tagsHeader.textContent = 'Tags:';
        tagsHeader.style.width = '100%';
        tagsHeader.style.fontSize = '0.75rem';
        tagsHeader.style.color = 'var(--spire-muted)';
        tagsWrapper.appendChild(tagsHeader);
        if (!Array.isArray(item.tags)) item.tags = [];
        WEAPON_TAGS.forEach(tag => {
          const pill = document.createElement('button');
          pill.type = 'button';
          pill.textContent = tag;
          pill.className = 'tag-toggle-pill' + (item.tags.includes(tag) ? ' active' : '');
          pill.addEventListener('click', () => {
            pushSectionUndo(pc, 'inventory', 'Edit weapon tags');
            if (item.tags.includes(tag)) {
              item.tags.splice(item.tags.indexOf(tag), 1);
            } else {
              item.tags.push(tag);
            }
            appendLog('Edited weapon tags', pc.id);
            saveAndRefresh();
          });
          tagsWrapper.appendChild(pill);
        });
        // Assemble
        const stressWrapper = document.createElement('div');
        stressWrapper.style.display = 'flex';
        stressWrapper.style.alignItems = 'center';
        stressWrapper.style.gap = '4px';
        stressWrapper.appendChild(stressLabel);
        stressWrapper.appendChild(stressSelect);
        typeFields.appendChild(stressWrapper);
        typeFields.appendChild(tagsWrapper);
      } else if (item.type === 'armor') {
        // Resistance input
        const resLabel = document.createElement('label');
        resLabel.textContent = 'Resistance:';
        const resInput = document.createElement('input');
        resInput.type = 'number';
        resInput.min = '1';
        resInput.max = '99';
        resInput.value = item.resistance || 0;
        resInput.style.width = '60px';
        resInput.addEventListener('change', e => {
          pushSectionUndo(pc, 'inventory', 'Edit armor resistance');
          item.resistance = parseInt(e.target.value) || 0;
          appendLog('Edited armor resistance', pc.id);
          saveAndRefresh();
        });
        // Armor tags — toggleable pills
        const tagsWrapper = document.createElement('div');
        tagsWrapper.style.display = 'flex';
        tagsWrapper.style.flexWrap = 'wrap';
        tagsWrapper.style.gap = '4px';
        tagsWrapper.style.marginTop = '6px';
        const tagsHeader = document.createElement('div');
        tagsHeader.textContent = 'Tags:';
        tagsHeader.style.width = '100%';
        tagsHeader.style.fontSize = '0.75rem';
        tagsHeader.style.color = 'var(--spire-muted)';
        tagsWrapper.appendChild(tagsHeader);
        if (!Array.isArray(item.tags)) item.tags = [];
        ARMOR_TAGS.forEach(tag => {
          const pill = document.createElement('button');
          pill.type = 'button';
          pill.textContent = tag;
          pill.className = 'tag-toggle-pill' + (item.tags.includes(tag) ? ' active' : '');
          pill.addEventListener('click', () => {
            pushSectionUndo(pc, 'inventory', 'Edit armor tags');
            if (item.tags.includes(tag)) {
              item.tags.splice(item.tags.indexOf(tag), 1);
            } else {
              item.tags.push(tag);
            }
            appendLog('Edited armor tags', pc.id);
            saveAndRefresh();
          });
          tagsWrapper.appendChild(pill);
        });
        // Assemble
        const resWrapper = document.createElement('div');
        resWrapper.style.display = 'flex';
        resWrapper.style.alignItems = 'center';
        resWrapper.style.gap = '4px';
        resWrapper.appendChild(resLabel);
        resWrapper.appendChild(resInput);
        typeFields.appendChild(resWrapper);
        typeFields.appendChild(tagsWrapper);
      } else {
        // Generic item: optional free-form tags (comma separated)
        const tagInput = document.createElement('input');
        tagInput.type = 'text';
        tagInput.placeholder = 'Tags (comma separated)';
        tagInput.value = Array.isArray(item.tags) ? item.tags.join(', ') : '';
        tagInput.style.flex = '2';
        tagInput.addEventListener('change', e => {
          pushSectionUndo(pc, 'inventory', 'Edit item tags');
          const val = e.target.value.trim();
          item.tags = val ? val.split(',').map(s => s.trim()).filter(Boolean) : [];
          appendLog('Edited item tags', pc.id);
          saveAndRefresh();
        });
        typeFields.appendChild(tagInput);
      }
    }
    // Initial render of type-specific fields
    renderTypeFields();

    // Remove button
    const remBtn = document.createElement('button');
    remBtn.textContent = '×';
    remBtn.className = 'row-remove-btn';
    remBtn.addEventListener('click', () => {
      pushSectionUndo(pc, 'inventory', 'Remove inventory item');
      pc.inventory = pc.inventory.filter(x => x.id !== item.id);
      appendLog('Removed inventory', pc.id);
      saveAndRefresh();
    });

    // When type changes, re-render specific fields
    typeSelect.addEventListener('change', () => {
      renderTypeFields();
    });

    // Append all elements
    row.appendChild(typeSelect);
    row.appendChild(itemInput);
    row.appendChild(qtyInput);
    row.appendChild(notesInput);
    row.appendChild(typeFields);
    row.appendChild(remBtn);
    return row;
  }

  /**
   * Helper to create a task/quest row. Allows editing title, status,
   * priority, due date and notes.
   */
  function createTaskRow(pc, task) {
    const row = document.createElement('div');
    row.className = 'task-row';
    row.style.display = 'flex';
    row.style.alignItems = 'center';
    row.style.marginBottom = '4px';
    const titleInput = document.createElement('input');
    titleInput.type = 'text';
    titleInput.placeholder = 'Title';
    titleInput.value = task.title || '';
    titleInput.style.flex = '2';
    titleInput.addEventListener('input', e => {
      task.title = e.target.value;
      queueDeferredSave(`task-title:${pc.id}:${task.id}`);
    });
    titleInput.addEventListener('change', e => {
      pushSectionUndo(pc, 'tasks', 'Edit task title');
      task.title = e.target.value;
      appendLog('Edited task', pc.id);
      saveWithoutRefresh();
    });
    const statusSelect = document.createElement('select');
    ['To Do','Doing','Done'].forEach(st => {
      const opt = document.createElement('option');
      opt.value = st;
      opt.textContent = st;
      if (task.status === st) opt.selected = true;
      statusSelect.appendChild(opt);
    });
    statusSelect.addEventListener('change', e => {
      pushSectionUndo(pc, 'tasks', 'Edit task status');
      task.status = e.target.value;
      appendLog('Edited task', pc.id);
      saveWithoutRefresh();
    });
    const prioritySelect = document.createElement('select');
    ['Low','Normal','High','Urgent'].forEach(pr => {
      const opt = document.createElement('option');
      opt.value = pr;
      opt.textContent = pr;
      if (task.priority === pr) opt.selected = true;
      prioritySelect.appendChild(opt);
    });
    prioritySelect.addEventListener('change', e => {
      pushSectionUndo(pc, 'tasks', 'Edit task priority');
      task.priority = e.target.value;
      appendLog('Edited task', pc.id);
      saveWithoutRefresh();
    });
    const dueInput = document.createElement('input');
    dueInput.type = 'date';
    dueInput.value = task.dueDate || '';
    dueInput.addEventListener('change', e => {
      pushSectionUndo(pc, 'tasks', 'Edit task due date');
      task.dueDate = e.target.value;
      appendLog('Edited task', pc.id);
      saveWithoutRefresh();
    });
    const notesInput = document.createElement('input');
    notesInput.type = 'text';
    notesInput.placeholder = 'Notes';
    notesInput.value = task.notes || '';
    notesInput.style.flex = '3';
    notesInput.addEventListener('input', e => {
      task.notes = e.target.value;
      queueDeferredSave(`task-notes:${pc.id}:${task.id}`);
    });
    notesInput.addEventListener('change', e => {
      pushSectionUndo(pc, 'tasks', 'Edit task notes');
      task.notes = e.target.value;
      appendLog('Edited task', pc.id);
      saveWithoutRefresh();
    });
    const remBtn = document.createElement('button');
    remBtn.textContent = '×';
    remBtn.className = 'row-remove-btn';
    remBtn.addEventListener('click', () => {
      pushSectionUndo(pc, 'tasks', 'Remove task');
      pc.tasks = pc.tasks.filter(x => x.id !== task.id);
      appendLog('Removed task', pc.id);
      saveAndRefresh();
    });
    row.appendChild(titleInput);
    row.appendChild(statusSelect);
    row.appendChild(prioritySelect);
    row.appendChild(dueInput);
    row.appendChild(notesInput);
    row.appendChild(remBtn);
    return row;
  }

  /**
   * Escape HTML for safe insertion into the DOM.
   */
  function escapeHtml(str) {
    return String(str || '').replace(/[&<>"']/g, function (m) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m];
    });
  }

  /**
   * Refresh UI after modifying data: update lists, sheet view and graph.
   */
  function showToast(message, type = 'info') {
    let toast = document.getElementById('spire-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'spire-toast';
      document.body.appendChild(toast);
    }
    toast.textContent = message;
    toast.className = 'spire-toast spire-toast-' + type + ' spire-toast-show';
    clearTimeout(toast._timer);
    toast._timer = setTimeout(() => {
      toast.classList.remove('spire-toast-show');
    }, 3500);
  }

  function setSaveState(mode = 'saved', message = '') {
    const el = document.getElementById('save-state-indicator');
    const retryBtn = document.getElementById('save-retry-btn');
    if (!el) return;
    if (saveStateTimer) {
      clearTimeout(saveStateTimer);
      saveStateTimer = null;
    }
    el.classList.remove('state-saving', 'state-saved', 'state-error');
    if (mode === 'saving') {
      el.classList.add('state-saving');
      el.textContent = message || 'Saving...';
      el.title = 'Saving changes to local storage';
      if (retryBtn) retryBtn.classList.add('hidden');
      return;
    }
    if (mode === 'error') {
      el.classList.add('state-error');
      el.textContent = message || 'Save failed';
      el.title = message || 'Save failed';
      if (retryBtn) retryBtn.classList.remove('hidden');
      return;
    }
    el.classList.add('state-saved');
    const stamp = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    el.textContent = message || `Saved ${stamp}`;
    el.title = `Last saved at ${stamp}`;
    if (retryBtn) retryBtn.classList.add('hidden');
    saveStateTimer = setTimeout(() => {
      el.classList.remove('state-saved');
      el.textContent = `Saved ${stamp}`;
    }, 1500);
  }

  function formatSyncOpSummary(op) {
    if (!op || !op.payload) return 'Queued change';
    const campCount = Object.keys(op.payload.campaigns || {}).length;
    const active = op.payload.currentCampaignId || 'none';
    return `${campCount} campaign${campCount === 1 ? '' : 's'} (active: ${active})`;
  }

  function updateSyncQueueButton() {
    const btn = document.getElementById('sync-queue-btn');
    if (!btn) return;
    const badge = document.getElementById('sync-queue-count');
    const count = state.currentUser ? loadPendingOps().length : 0;
    btn.title = count ? `Sync queue (${count} pending)` : 'Sync queue (empty)';
    btn.classList.toggle('has-pending', count > 0);
    if (badge) {
      badge.textContent = String(count);
      badge.classList.toggle('hidden', count <= 0);
    }
  }

  function setSyncConflictWarning(show, message = 'Updated in another tab') {
    const el = document.getElementById('sync-conflict-indicator');
    const reloadBtn = document.getElementById('sync-reload-btn');
    const forceBtn = document.getElementById('sync-force-btn');
    state.syncConflictActive = !!show;
    if (!show) state.localEditsSinceConflict = 0;
    if (el) {
      el.textContent = message;
      el.title = message;
      el.classList.toggle('hidden', !show);
    }
    if (reloadBtn) reloadBtn.classList.toggle('hidden', !show);
    if (forceBtn) forceBtn.classList.toggle('hidden', !show);
  }

  function forceOverwriteSave() {
    const saved = saveCampaigns({ force: true });
    if (saved) {
      showToast('Campaign saved (forced overwrite).', 'warn');
    } else {
      showToast('Force save failed.', 'warn');
    }
  }

  async function attemptManualSave() {
    if (!state.syncConflictActive) {
      const saved = saveCampaigns();
      if (saved) showToast('Campaign saved', 'info');
      return;
    }
    const ok = await askConfirm(
      'This campaign changed in another tab. Save anyway and overwrite localStorage?',
      'Sync Conflict'
    );
    if (!ok) return;
    const saved = saveCampaigns({ force: true });
    if (saved) showToast('Campaign saved (forced overwrite).', 'warn');
  }

  function reloadCampaignFromStorage() {
    if (!state.currentUser) return;
    try {
      const raw = localStorage.getItem(userScopedKey('spire-campaigns'));
      if (!raw) {
        showToast('No stored campaign data to reload.', 'warn');
        return;
      }
      const parsed = JSON.parse(raw);
      state.campaigns = parsed.campaigns || {};
      state.currentCampaignId = parsed.currentCampaignId;
      if (!state.currentCampaignId || !state.campaigns[state.currentCampaignId]) {
        const first = Object.keys(state.campaigns)[0];
        state.currentCampaignId = first || null;
      }
      if (!state.currentCampaignId) {
        const camp = createCampaign('Default');
        state.campaigns[camp.id] = camp;
        state.currentCampaignId = camp.id;
      }
      state.lastSeenRevision = localStorage.getItem(userScopedKey('spire-campaigns-rev')) || state.lastSeenRevision;
      initAfterLoad();
      setSyncConflictWarning(false);
      showToast('Reloaded latest campaign data from storage.', 'info');
    } catch (e) {
      console.warn('Failed to reload from storage', e);
      showToast('Could not reload latest data.', 'warn');
    }
  }

  function openSyncConflictModal() {
    if (!state.syncConflictActive) return;
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    const content = document.getElementById('modal-content');
    const titleEl = document.getElementById('modal-title');
    if (titleEl) titleEl.textContent = 'Sync Conflict';
    content.innerHTML = '';

    const msg = document.createElement('p');
    msg.style.marginBottom = '12px';
    msg.textContent = `${conflictWarningText()}. Choose how to resolve:`;
    content.appendChild(msg);

    const actions = document.createElement('div');
    actions.className = 'modal-form';
    actions.style.display = 'flex';
    actions.style.flexDirection = 'column';
    actions.style.gap = '8px';

    const reloadBtn = document.createElement('button');
    reloadBtn.className = 'modal-submit';
    reloadBtn.textContent = 'Reload Latest From Other Tab';
    reloadBtn.addEventListener('click', () => {
      closeModal();
      reloadCampaignFromStorage();
    });
    actions.appendChild(reloadBtn);

    const keepBtn = document.createElement('button');
    keepBtn.className = 'modal-submit';
    keepBtn.style.background = 'var(--spire-mid)';
    keepBtn.textContent = 'Keep Mine (Force Save)';
    keepBtn.addEventListener('click', () => {
      closeModal();
      const saved = saveCampaigns({ force: true });
      if (saved) showToast('Campaign saved (forced overwrite).', 'warn');
    });
    actions.appendChild(keepBtn);

    const dismissBtn = document.createElement('button');
    dismissBtn.className = 'modal-submit';
    dismissBtn.style.background = 'transparent';
    dismissBtn.textContent = 'Dismiss';
    dismissBtn.addEventListener('click', () => closeModal());
    actions.appendChild(dismissBtn);

    content.appendChild(actions);
    overlay.classList.remove('hidden');
    modal.classList.remove('hidden');
  }

  function openSyncQueueModal() {
    if (!state.currentUser) return;
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    const content = document.getElementById('modal-content');
    const titleEl = document.getElementById('modal-title');
    const ops = loadPendingOps();
    if (titleEl) titleEl.textContent = `Sync Queue (${ops.length})`;
    content.innerHTML = '';
    overlay.classList.remove('hidden');
    modal.classList.remove('hidden');

    const hint = document.createElement('p');
    hint.className = 'text-muted';
    hint.style.marginBottom = '10px';
    hint.textContent = ops.length
      ? 'Queued saves were created while offline or blocked by conflict. Retry sync or discard stale items.'
      : 'No queued save operations.';
    content.appendChild(hint);

    const list = document.createElement('div');
    list.className = 'sync-queue-list';
    const selectedIds = new Set();
    ops.forEach((op) => {
      const row = document.createElement('label');
      row.className = 'sync-queue-row';
      const chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.addEventListener('change', (e) => {
        if (e.target.checked) selectedIds.add(op.id);
        else selectedIds.delete(op.id);
      });
      const body = document.createElement('div');
      body.className = 'sync-queue-row-body';
      const when = document.createElement('div');
      when.className = 'sync-queue-when';
      const ts = op.time ? new Date(op.time) : null;
      when.textContent = ts && !Number.isNaN(ts.getTime()) ? ts.toLocaleString() : 'Unknown time';
      const meta = document.createElement('div');
      meta.className = 'sync-queue-meta';
      const baseRev = op.baseRevision ? `base ${String(op.baseRevision).slice(0, 12)}...` : 'base none';
      meta.textContent = `${op.kind || 'op'} • ${baseRev}`;
      const summary = document.createElement('div');
      summary.className = 'sync-queue-summary';
      summary.textContent = formatSyncOpSummary(op);
      body.appendChild(when);
      body.appendChild(meta);
      body.appendChild(summary);
      row.appendChild(chk);
      row.appendChild(body);
      list.appendChild(row);
    });
    if (ops.length) content.appendChild(list);

    const actions = document.createElement('div');
    actions.className = 'modal-form';
    actions.style.display = 'flex';
    actions.style.gap = '8px';
    actions.style.marginTop = '10px';
    actions.style.flexWrap = 'wrap';

    const retryBtn = document.createElement('button');
    retryBtn.className = 'modal-submit';
    retryBtn.textContent = 'Retry Sync';
    retryBtn.disabled = !ops.length;
    retryBtn.addEventListener('click', () => {
      const ok = flushPendingSaveOps();
      if (!ok) {
        showToast(navigator.onLine ? 'Sync failed. Resolve conflict or discard queued ops.' : 'Still offline. Queue unchanged.', 'warn');
      } else {
        showToast('Queued changes synced.', 'info');
      }
      openSyncQueueModal();
    });
    actions.appendChild(retryBtn);

    const discardSelectedBtn = document.createElement('button');
    discardSelectedBtn.textContent = 'Discard Selected';
    discardSelectedBtn.disabled = !ops.length;
    discardSelectedBtn.addEventListener('click', async () => {
      if (!selectedIds.size) {
        showToast('No queued items selected.', 'warn');
        return;
      }
      const ok = await askConfirm(`Discard ${selectedIds.size} queued item(s)?`, 'Discard Queued Items');
      if (!ok) return;
      removePendingOpsById(Array.from(selectedIds));
      showToast('Selected queued items discarded.', 'info');
      openSyncQueueModal();
    });
    actions.appendChild(discardSelectedBtn);

    const discardAllBtn = document.createElement('button');
    discardAllBtn.className = 'toolbar-btn-danger';
    discardAllBtn.textContent = 'Discard All';
    discardAllBtn.disabled = !ops.length;
    discardAllBtn.addEventListener('click', async () => {
      const ok = await askConfirm('Discard all queued save operations?', 'Discard Queue');
      if (!ok) return;
      savePendingOps([]);
      showToast('Sync queue cleared.', 'warn');
      openSyncQueueModal();
    });
    actions.appendChild(discardAllBtn);
    content.appendChild(actions);
  }

  let modalDecisionResolver = null;

  function resolveModalDecision(value) {
    if (modalDecisionResolver) {
      const resolver = modalDecisionResolver;
      modalDecisionResolver = null;
      resolver(value);
    }
  }

  function askConfirm(message, title = 'Confirm') {
    return new Promise((resolve) => {
      const overlay = document.getElementById('modal-overlay');
      const modal = document.getElementById('modal');
      const content = document.getElementById('modal-content');
      const titleEl = document.getElementById('modal-title');
      if (titleEl) titleEl.textContent = title;
      content.innerHTML = '';

      const msg = document.createElement('p');
      msg.style.margin = '0 0 12px';
      msg.textContent = message;
      content.appendChild(msg);

      const row = document.createElement('div');
      row.style.display = 'flex';
      row.style.justifyContent = 'flex-end';
      row.style.gap = '8px';
      const cancelBtn = document.createElement('button');
      cancelBtn.textContent = 'Cancel';
      const okBtn = document.createElement('button');
      okBtn.className = 'modal-submit';
      okBtn.textContent = 'Confirm';
      row.appendChild(cancelBtn);
      row.appendChild(okBtn);
      content.appendChild(row);

      modalDecisionResolver = resolve;
      cancelBtn.addEventListener('click', () => {
        closeModal(false);
      });
      okBtn.addEventListener('click', () => {
        closeModal(true);
      });
      overlay.classList.remove('hidden');
      modal.classList.remove('hidden');
    });
  }

  function askPrompt(message, defaultValue = '', opts = {}) {
    return new Promise((resolve) => {
      const overlay = document.getElementById('modal-overlay');
      const modal = document.getElementById('modal');
      const content = document.getElementById('modal-content');
      const titleEl = document.getElementById('modal-title');
      if (titleEl) titleEl.textContent = opts.title || 'Input';
      content.innerHTML = '';

      const msg = document.createElement('p');
      msg.style.margin = '0 0 10px';
      msg.textContent = message;
      content.appendChild(msg);

      const input = document.createElement('input');
      input.type = opts.type || 'text';
      input.placeholder = opts.placeholder || '';
      input.value = defaultValue || '';
      input.style.width = '100%';
      input.style.marginBottom = '10px';
      content.appendChild(input);

      const row = document.createElement('div');
      row.style.display = 'flex';
      row.style.justifyContent = 'flex-end';
      row.style.gap = '8px';
      const cancelBtn = document.createElement('button');
      cancelBtn.textContent = 'Cancel';
      const okBtn = document.createElement('button');
      okBtn.className = 'modal-submit';
      okBtn.textContent = opts.submitText || 'Save';
      row.appendChild(cancelBtn);
      row.appendChild(okBtn);
      content.appendChild(row);

      modalDecisionResolver = resolve;
      cancelBtn.addEventListener('click', () => {
        closeModal(null);
      });
      okBtn.addEventListener('click', () => {
        const val = input.value;
        closeModal(val);
      });
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') okBtn.click();
      });
      overlay.classList.remove('hidden');
      modal.classList.remove('hidden');
      input.focus();
      input.select();
    });
  }

  function updatePendingBadge() {
    const camp = currentCampaign();
    const count = Object.values(camp.entities).filter(e => e.pendingApproval).length;
    const badge = document.getElementById('pending-badge');
    if (!badge) return;
    if (count > 0 && state.gmMode) {
      badge.textContent = count + ' pending';
      badge.classList.remove('hidden');
    } else {
      badge.classList.add('hidden');
    }
  }

  function updatePlayerPCButtonState() {
    const btn = document.getElementById('player-add-pc-btn');
    if (!btn) return;
    const camp = currentCampaign();
    const ownedId = camp.playerOwnedPcId;
    const ownedPc = ownedId ? camp.entities[ownedId] : null;
    const anyPcExists = Object.values(camp.entities).some(e => e.type === 'pc');
    btn.disabled = !state.gmMode && !!ownedPc;
    const label = btn.querySelector('span:last-child');
    if (ownedPc) {
      btn.title = 'You already claimed a character';
      if (label) label.textContent = 'My Character';
      return;
    }
    btn.title = anyPcExists ? 'Claim an existing character or create a new one' : 'Create my character';
    if (label) label.textContent = anyPcExists ? 'Claim / Create' : 'My Character';
  }

  function updateFocusToggleUI() {
    const btn = document.getElementById('filter-focus-toggle');
    if (!btn) return;
    btn.textContent = 'Focus mode: ' + (graphState.focusMode ? 'On' : 'Off');
    btn.classList.toggle('active', graphState.focusMode);
  }

  function setGraphFocusMode(enabled) {
    graphState.focusMode = !!enabled;
    updateFocusToggleUI();
    drawGraph();
  }

  function getFocusedGraphNodeId() {
    if (!graphState.focusMode) return null;
    if (state.selectedEntityId && graphNodes.some(n => n.id === state.selectedEntityId)) return state.selectedEntityId;
    if (graphSelectedNodeId && graphNodes.some(n => n.id === graphSelectedNodeId)) return graphSelectedNodeId;
    return null;
  }

  function makeInteractivePip(pipEl, label) {
    pipEl.setAttribute('role', 'button');
    pipEl.setAttribute('tabindex', '0');
    pipEl.setAttribute('aria-label', label);
    pipEl.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        pipEl.click();
      }
    });
  }

  function openClaimPCModal(openCreateFlow) {
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    const content = document.getElementById('modal-content');
    const titleEl = document.getElementById('modal-title');
    if (titleEl) titleEl.textContent = 'Choose Character';
    content.innerHTML = '';
    overlay.classList.remove('hidden');
    modal.classList.remove('hidden');

    const camp = currentCampaign();
    const pcs = Object.values(camp.entities).filter(e => e.type === 'pc');
    const intro = document.createElement('p');
    intro.textContent = 'Claim an existing character, or create a new one.';
    intro.style.color = 'var(--spire-muted)';
    intro.style.marginBottom = '10px';
    content.appendChild(intro);

    const field = document.createElement('div');
    field.className = 'modal-field';
    const label = document.createElement('label');
    label.textContent = 'Existing Characters';
    const sel = document.createElement('select');
    pcs.forEach(pc => {
      const opt = document.createElement('option');
      opt.value = pc.id;
      opt.textContent = entityLabel(pc);
      sel.appendChild(opt);
    });
    field.appendChild(label);
    field.appendChild(sel);
    content.appendChild(field);

    const btnBar = document.createElement('div');
    btnBar.style.display = 'flex';
    btnBar.style.gap = '8px';
    btnBar.style.marginTop = '10px';

    const claimBtn = document.createElement('button');
    claimBtn.className = 'modal-submit';
    claimBtn.textContent = 'Claim Selected';
    claimBtn.addEventListener('click', () => {
      const chosen = camp.entities[sel.value];
      if (!chosen) return;
      camp.playerOwnedPcId = chosen.id;
      closeModal();
      saveAndRefresh();
      selectEntity(chosen.id);
      showToast('Character claimed: ' + entityLabel(chosen), 'info');
    });

    const createBtn = document.createElement('button');
    createBtn.textContent = 'Create New Instead';
    createBtn.addEventListener('click', () => {
      closeModal();
      openCreateFlow();
    });

    btnBar.appendChild(claimBtn);
    btnBar.appendChild(createBtn);
    content.appendChild(btnBar);
  }

  function totalStressForFallout(pc) {
    const tracks = ['blood', 'mind', 'silver', 'shadow', 'reputation'];
    return tracks.reduce((sum, track) => {
      const filled = (pc && pc.stressFilled && pc.stressFilled[track]) ? pc.stressFilled[track].length : 0;
      let free = 0;
      if (Array.isArray(pc && pc.resistances)) {
        pc.resistances.forEach((r) => {
          if (!r || !r.name) return;
          if (String(r.name).toLowerCase() !== track) return;
          free += Math.max(0, parseInt(r.value, 10) || 0);
        });
      }
      if (track === 'blood' && Array.isArray(pc && pc.inventory)) {
        pc.inventory.forEach((item) => {
          if (!item || item.type !== 'armor') return;
          free += Math.max(0, parseInt(item.resistance, 10) || 0);
        });
      }
      return sum + Math.min(10, Math.max(0, filled - free));
    }, 0);
  }

  function falloutSeverityForTotalStress(total) {
    return RulesEngine.falloutSeverityForTotalStress(total);
  }

  function stressClearAmountForSeverity(severity) {
    return RulesEngine.stressClearAmountForSeverity(severity);
  }

  function clearStressForFallout(pc, amount, preferredTrack) {
    const tracks = ['blood', 'mind', 'silver', 'shadow', 'reputation'];
    let remaining = amount;
    const ordered = tracks.slice();
    if (preferredTrack && tracks.includes(preferredTrack)) {
      ordered.splice(ordered.indexOf(preferredTrack), 1);
      ordered.unshift(preferredTrack);
    }
    while (remaining > 0) {
      const track = ordered.find(t => (pc.stressFilled[t] || []).length > 0);
      if (!track) break;
      pc.stressFilled[track].pop();
      remaining--;
    }
  }

  function setPCStressLevel(pc, track, newLevel, options = {}) {
    const triggerFallout = options.triggerFallout !== false;
    const rng = options.rng;
    const slots = getTrackTotalSlots(pc, track);
    const target = Math.max(0, Math.min(slots, parseInt(newLevel, 10) || 0));
    const before = (pc.stressFilled && pc.stressFilled[track]) ? pc.stressFilled[track].length : 0;
    pc.stressFilled[track] = Array.from({ length: target }, (_, i) => i);
    appendLog('Set ' + track + ' stress to ' + target, pc.id);
    const delta = target - before;
    if (triggerFallout) maybeTriggerFallout(pc, track, delta, { rng });
  }

  function maybeTriggerFallout(pc, contextTrack, stressDelta, options = {}) {
    const rules = getRulesConfig();
    if (!rules.falloutCheckOnStress) return;
    if (stressDelta <= 0) return;
    const total = totalStressForFallout(pc);
    if (total < 2) return;
    const rng = (options && typeof options.rng === 'function') ? options.rng : Math.random;
    const roll = Math.ceil(rng() * 10);
    if (roll >= total) return;

    const severity = falloutSeverityForTotalStress(total);
    if (rules.clearStressOnFallout) {
      const stressToClear = stressClearAmountForSeverity(severity);
      clearStressForFallout(pc, stressToClear, contextTrack);
    }
    const fallout = {
      id: generateId('fallout'),
      type: contextTrack ? contextTrack.charAt(0).toUpperCase() + contextTrack.slice(1) : 'General',
      severity,
      name: '',
      description: `Auto-triggered from stress check (D10 ${roll} vs total ${total}).`,
      resolved: false,
      timestamp: new Date().toISOString()
    };
    pc.fallout.push(fallout);
    appendLog(`Fallout triggered (${severity})`, pc.id);
    showToast(`Fallout triggered (${severity}) — rolled ${roll} vs total stress ${total}.`, 'warn');
  }

  function saveAndRefresh(options = {}) {
    const activeTab = document.querySelector('.tab-link.active')?.dataset?.tab || 'sheets-view';
    const forceSheetRender = !!options.forceSheetRender;
    const forceGraphRender = !!options.forceGraphRender;
    const forceMessagesRender = !!options.forceMessagesRender;
    const sheetBefore = document.getElementById('sheet-container');
    const preservedSheetScroll = sheetBefore && sheetBefore.style.display !== 'none' ? sheetBefore.scrollTop : null;
    const preservedScrolls = {
      sidebar: document.getElementById('sidebar')?.scrollTop || 0,
      logList: document.getElementById('log-list')?.scrollTop || 0,
      messagesList: document.getElementById('messages-list')?.scrollTop || 0,
      prep: document.getElementById('session-prep-panel')?.scrollTop || 0,
      gmNotes: document.getElementById('gm-notes-board')?.scrollTop || 0
    };
    saveCampaigns();
    renderEntityLists();
    if (forceSheetRender || activeTab === 'sheets-view') {
      renderSheetView();
      if (preservedSheetScroll !== null) {
        const sheetAfter = document.getElementById('sheet-container');
        if (sheetAfter) sheetAfter.scrollTop = preservedSheetScroll;
      }
    }
    if (forceGraphRender || activeTab === 'web-view') updateGraph();
    if (forceMessagesRender || activeTab === 'messages-view') renderMessages();
    if (activeTab === 'log-view') renderLog();
    if (activeTab === 'gm-notes-view') renderGMNotes();
    const sidebarAfter = document.getElementById('sidebar');
    if (sidebarAfter) sidebarAfter.scrollTop = preservedScrolls.sidebar;
    const logAfter = document.getElementById('log-list');
    if (logAfter) logAfter.scrollTop = preservedScrolls.logList;
    const msgAfter = document.getElementById('messages-list');
    if (msgAfter) msgAfter.scrollTop = preservedScrolls.messagesList;
    const prepAfter = document.getElementById('session-prep-panel');
    if (prepAfter) prepAfter.scrollTop = preservedScrolls.prep;
    const gmNotesAfter = document.getElementById('gm-notes-board');
    if (gmNotesAfter) gmNotesAfter.scrollTop = preservedScrolls.gmNotes;
    updatePendingBadge();
    updatePlayerPCButtonState();
    updateFocusToggleUI();
    updateMessagesUnreadBadge();
    updateUndoButtonState();
    ensureAutoGrowTextareas(document);
    ensureAccessibilityLabels(document);
    ensureAccessibilitySemantics();
  }

  function saveWithoutRefresh() {
    saveCampaigns();
    updatePendingBadge();
    updatePlayerPCButtonState();
    updateFocusToggleUI();
    updateMessagesUnreadBadge();
    updateUndoButtonState();
    ensureAutoGrowTextareas(document);
    ensureAccessibilityLabels(document);
    ensureAccessibilitySemantics();
    flashSavedSection();
  }

  function queueDeferredSave(key, delayMs = 280) {
    if (!key) {
      saveWithoutRefresh();
      return;
    }
    if (deferredSaveTimers.has(key)) clearTimeout(deferredSaveTimers.get(key));
    const timer = setTimeout(() => {
      deferredSaveTimers.delete(key);
      saveWithoutRefresh();
    }, delayMs);
    deferredSaveTimers.set(key, timer);
  }

  function flashSavedSection(sourceEl = null) {
    const active = sourceEl || document.activeElement;
    if (!active || !active.closest) return;
    const section = active.closest('.inspector-section');
    if (!section) return;
    const stampHost = section.querySelector('h3') || section.querySelector('.section-header');
    if (stampHost) {
      let stamp = stampHost.querySelector('.section-saved-stamp');
      if (!stamp) {
        stamp = document.createElement('span');
        stamp.className = 'section-saved-stamp';
        stampHost.appendChild(stamp);
      }
      stamp.textContent = 'Saved ' + new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    }
    section.classList.remove('saved-pulse');
    // Trigger reflow so repeated pulses still animate.
    void section.offsetWidth;
    section.classList.add('saved-pulse');
    setTimeout(() => section.classList.remove('saved-pulse'), 520);
  }

  function ensureAutoGrowTextareas(root = document) {
    if (!root || !root.querySelectorAll) return;
    root.querySelectorAll('textarea').forEach((ta) => {
      if (ta.dataset.autogrowBound === '1') return;
      ta.dataset.autogrowBound = '1';
      ta.classList.add('auto-grow');
      const resize = () => {
        ta.style.height = 'auto';
        ta.style.height = `${Math.max(ta.scrollHeight, 64)}px`;
      };
      ta.addEventListener('input', resize);
      requestAnimationFrame(resize);
    });
  }

  function ensureAccessibilityLabels(root = document) {
    if (!root || !root.querySelectorAll) return;
    const controls = root.querySelectorAll('button, input, select, textarea');
    controls.forEach((el) => {
      if (el.hasAttribute('aria-label')) return;
      if (el.id) {
        const direct = root.querySelector(`label[for="${el.id}"]`) || document.querySelector(`label[for="${el.id}"]`);
        if (direct && direct.textContent && direct.textContent.trim()) {
          el.setAttribute('aria-label', direct.textContent.trim());
          return;
        }
      }
      const title = (el.getAttribute('title') || '').trim();
      if (title) {
        el.setAttribute('aria-label', title);
        return;
      }
      const placeholder = (el.getAttribute('placeholder') || '').trim();
      if (placeholder) {
        el.setAttribute('aria-label', placeholder);
        return;
      }
      if (el.tagName === 'BUTTON') {
        const text = (el.textContent || '').trim();
        if (text) el.setAttribute('aria-label', text);
      }
    });
  }

  function ensureAccessibilitySemantics() {
    const tabWrap = document.querySelector('#top-nav .nav-middle');
    if (tabWrap) tabWrap.setAttribute('role', 'tablist');
    document.querySelectorAll('.tab-link').forEach((btn) => {
      const targetId = btn.dataset.tab;
      if (!targetId) return;
      const isActive = btn.classList.contains('active');
      btn.setAttribute('role', 'tab');
      btn.setAttribute('aria-controls', targetId);
      btn.setAttribute('aria-selected', isActive ? 'true' : 'false');
      btn.setAttribute('tabindex', isActive ? '0' : '-1');
      const panel = document.getElementById(targetId);
      if (panel) {
        panel.setAttribute('role', 'tabpanel');
        panel.setAttribute('aria-hidden', isActive ? 'false' : 'true');
      }
    });
    const modal = document.getElementById('modal');
    const modalTitle = document.getElementById('modal-title');
    if (modal) {
      modal.setAttribute('role', 'dialog');
      modal.setAttribute('aria-modal', 'true');
      if (modalTitle && modalTitle.id) modal.setAttribute('aria-labelledby', modalTitle.id);
    }
  }

  function ensureGraphViews(camp = currentCampaign()) {
    if (!camp.graphViews) camp.graphViews = {};
  }

  function syncWebFilterPills() {
    document.querySelectorAll('#web-filter-bar .filter-pill[data-type]').forEach(pill => {
      const type = pill.dataset.type;
      pill.classList.toggle('active', !!webFilter[type]);
    });
  }

  function populateGraphViewSelect() {
    const sel = document.getElementById('graph-view-select');
    if (!sel) return;
    const camp = currentCampaign();
    ensureGraphViews(camp);
    const prior = sel.value;
    sel.innerHTML = '';
    const base = document.createElement('option');
    base.value = '';
    base.textContent = 'View preset…';
    sel.appendChild(base);
    Object.keys(camp.graphViews).sort((a, b) => a.localeCompare(b)).forEach(name => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      if (prior === name) opt.selected = true;
      sel.appendChild(opt);
    });
  }

  function saveCurrentGraphView(name) {
    const camp = currentCampaign();
    ensureGraphViews(camp);
    camp.graphViews[name] = {
      scale: graphState.scale,
      translateX: graphState.translateX,
      translateY: graphState.translateY,
      filters: Object.assign({}, webFilter)
    };
    saveCampaigns();
    populateGraphViewSelect();
    const sel = document.getElementById('graph-view-select');
    if (sel) sel.value = name;
    showToast('Saved graph view "' + name + '".', 'info');
  }

  function applyGraphView(name) {
    const camp = currentCampaign();
    ensureGraphViews(camp);
    const view = camp.graphViews[name];
    if (!view) return;
    if (view.filters) {
      Object.keys(webFilter).forEach(k => {
        if (Object.prototype.hasOwnProperty.call(view.filters, k)) webFilter[k] = !!view.filters[k];
      });
      syncWebFilterPills();
    }
    if (typeof view.scale === 'number') graphState.scale = Math.max(0.2, Math.min(4, view.scale));
    if (typeof view.translateX === 'number') graphState.translateX = view.translateX;
    if (typeof view.translateY === 'number') graphState.translateY = view.translateY;
    updateGraph();
    showToast('Loaded graph view "' + name + '".', 'info');
  }

  /**
   * Populate the node selection dropdown in the web toolbar. This allows
   * players and GMs to quickly select an entity in the conspiracy web
   * without needing to click its node directly. The select is cleared
   * and rebuilt each time the graph is updated or the entity list changes.
   */
  function populateNodeSelect() {
    const sel = document.getElementById('node-select');
    if (!sel) return;
    sel.innerHTML = '';
    const defaultOpt = document.createElement('option');
    defaultOpt.value = '';
    defaultOpt.textContent = '-- Select Entity --';
    sel.appendChild(defaultOpt);
    const camp = currentCampaign();
    Object.values(camp.entities).forEach(ent => {
      // Hide GM-only entities when not in GM mode
      if (!state.gmMode && ent.gmOnly) return;
      const opt = document.createElement('option');
      opt.value = ent.id;
      opt.textContent = entityLabel(ent);
      if (state.selectedEntityId === ent.id) opt.selected = true;
      sel.appendChild(opt);
    });
  }

  /**
   * Show a modal form for adding or editing a relationship. If
   * defaultSourceId is provided, the source select is prepopulated and
   * disabled. Otherwise the user can choose any source. The modal
   * includes fields for source, target, type and secrecy. On
   * confirmation, a new relationship is created and the modal closed.
   */
  function openRelModal(defaultSourceId) {
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    const content = document.getElementById('modal-content');
    // Set modal header title
    const titleEl = document.getElementById('modal-title');
    if (titleEl) titleEl.textContent = 'Add Relationship';
    // Clear previous content
    content.innerHTML = '';
    const camp = currentCampaign();
    // Form container
    const form = document.createElement('div');
    form.className = 'modal-form';
    // Source field
    const srcField = document.createElement('div');
    srcField.className = 'modal-field';
    const srcLabel = document.createElement('label');
    srcLabel.textContent = 'Source';
    const srcSelect = document.createElement('select');
    Object.values(camp.entities).forEach(ent => {
      if (!state.gmMode && ent.gmOnly) return;
      const opt = document.createElement('option');
      opt.value = ent.id;
      opt.textContent = entityLabel(ent);
      srcSelect.appendChild(opt);
    });
    if (defaultSourceId) {
      srcSelect.value = defaultSourceId;
      srcSelect.disabled = true;
    }
    srcField.appendChild(srcLabel);
    srcField.appendChild(srcSelect);
    form.appendChild(srcField);
    // Target field
    const tgtField = document.createElement('div');
    tgtField.className = 'modal-field';
    const tgtLabel = document.createElement('label');
    tgtLabel.textContent = 'Target';
    const tgtSelect = document.createElement('select');
    Object.values(camp.entities).forEach(ent => {
      if (!state.gmMode && ent.gmOnly) return;
      const opt = document.createElement('option');
      opt.value = ent.id;
      opt.textContent = entityLabel(ent);
      tgtSelect.appendChild(opt);
    });
    if (defaultSourceId) {
      // Do not pre-select the same entity as target
      if (tgtSelect.value === defaultSourceId) tgtSelect.value = '';
    }
    tgtField.appendChild(tgtLabel);
    tgtField.appendChild(tgtSelect);
    form.appendChild(tgtField);
    // Type field
    const typeField = document.createElement('div');
    typeField.className = 'modal-field';
    const typeLabel = document.createElement('label');
    typeLabel.textContent = 'Type';
    const typeSelect = document.createElement('select');
    state.relTypes.forEach(t => {
      const opt = document.createElement('option');
      opt.value = t;
      opt.textContent = t;
      typeSelect.appendChild(opt);
    });
    typeField.appendChild(typeLabel);
    typeField.appendChild(typeSelect);
    form.appendChild(typeField);
    // Secret checkbox
    const secretField = document.createElement('div');
    secretField.className = 'modal-field';
    const secretLabel = document.createElement('label');
    const secretChk = document.createElement('input');
    secretChk.type = 'checkbox';
    secretLabel.appendChild(secretChk);
    secretLabel.appendChild(document.createTextNode(' Secret (GM only)'));
    secretField.appendChild(secretLabel);
    form.appendChild(secretField);
    // Buttons
    const btnBar = document.createElement('div');
    btnBar.style.marginTop = '8px';
    btnBar.style.display = 'flex';
    btnBar.style.gap = '8px';
    const okBtn = document.createElement('button');
    okBtn.textContent = 'Add';
    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    btnBar.appendChild(okBtn);
    btnBar.appendChild(cancelBtn);
    form.appendChild(btnBar);
    // Append form
    content.appendChild(form);
    // Show modal
    overlay.classList.remove('hidden');
    modal.classList.remove('hidden');
    // Confirm handler
    okBtn.addEventListener('click', () => {
      const src = defaultSourceId || srcSelect.value;
      const tgt = tgtSelect.value;
      const type = typeSelect.value;
      const secret = secretChk.checked;
      // Validate: cannot link to self
      if (!src || !tgt) {
        showToast('Source and target must be selected.', 'warn');
        return;
      }
      if (src === tgt) {
        showToast('Source and target cannot be the same entity.', 'warn');
        return;
      }
      newRelationship(camp, src, tgt, type, { secret });
      appendLog('Added relationship', src);
      closeModal();
      saveAndRefresh();
    });
    // Cancel handler
    cancelBtn.addEventListener('click', () => {
      closeModal();
    });
  }

  /**
   * Close the generic modal. Hides overlay and modal elements.
   */
  function closeModal(decisionValue = null) {
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    overlay.classList.add('hidden');
    modal.classList.add('hidden');
    resolveModalDecision(decisionValue);
  }


  /**
   * Initialise Cytoscape and set up the graph view. The graph is
   * reconstructed whenever data changes. Nodes are styled based on
   * entity type and edges reflect relationship properties. Secret
   * edges and GM-only nodes are hidden in Player view.
   */
  function setupGraph() {
    // Initialise custom canvas-based graph. Obtain the canvas and context.
    graphCanvas = document.getElementById('graph-canvas');
    if (!graphCanvas) return;
    graphCtx = graphCanvas.getContext('2d');
    // Ensure the canvas fills its container
    resizeGraphCanvas();
    window.addEventListener('resize', resizeGraphCanvas);
    // Mouse and wheel event handlers for panning, zooming, dragging and
    // selection. Passive:false is used on wheel to prevent page scroll.
    graphCanvas.addEventListener('mousedown', onGraphMouseDown);
    graphCanvas.addEventListener('mousemove', onGraphMouseMove);
    graphCanvas.addEventListener('mouseup', onGraphMouseUp);
    graphCanvas.addEventListener('mouseleave', onGraphMouseUp);
    graphCanvas.addEventListener('wheel', onGraphWheel, { passive: false });
    graphCanvas.addEventListener('click', onGraphClick);
    graphCanvas.addEventListener('dblclick', onGraphDblClick);
    graphCanvas.addEventListener('contextmenu', onGraphContextMenu);
    // Build initial graph data and draw
    updateGraph();
  }

  /**
   * Rebuild the graph from campaign data. Applies classes to hide nodes
   * or edges based on GM mode and secrets. Uses preset positions when
   * available, otherwise runs an automatic layout.
   */
  // Web filter state: which entity types to show
  const webFilter = { pc: true, npc: true, org: true, secrets: true };

  function updateWebLegendMeta(camp, stats = {}) {
    const lineEl = document.getElementById('legend-semantic-line');
    const visEl = document.getElementById('legend-semantic-visibility');
    const relHint = document.getElementById('web-rel-hint');
    if (!camp) return;
    if (lineEl) {
      lineEl.textContent = `Line width: default 2px, selected 3px. Directed links show arrowheads. Focus mode ${graphState.focusMode ? 'de-emphasizes non-neighbors' : 'is off'}.`;
    }
    if (visEl) {
      const hiddenEntities = Math.max(0, (stats.totalEntities || 0) - (stats.visibleEntities || 0));
      const hiddenRelationships = Math.max(0, (stats.totalRelationships || 0) - (stats.visibleRelationships || 0));
      visEl.textContent = `Visibility: showing ${stats.visibleEntities || 0}/${stats.totalEntities || 0} entities and ${stats.visibleRelationships || 0}/${stats.totalRelationships || 0} relationships. Hidden now: ${hiddenEntities} entities, ${hiddenRelationships} relationships.`;
    }
    if (relHint) {
      relHint.textContent = (stats.visibleRelationships || 0) > 0
        ? `Current web: ${stats.visibleRelationships} visible links. Tip: double-click a node to open its sheet.`
        : 'Tip: add relationships to show edges between entities.';
    }
  }

  function updateGraph() {
    const camp = currentCampaign();
    graphNodes = [];
    graphEdges = [];
    const nodeMap = {};
    // Colours for different entity types, pulled from CSS variables
    const pcColor = getComputedStyle(document.documentElement).getPropertyValue('--pc-color').trim();
    const npcColor = getComputedStyle(document.documentElement).getPropertyValue('--npc-color').trim();
    const orgColor = getComputedStyle(document.documentElement).getPropertyValue('--org-color').trim();
    const entityPool = Object.values(camp.entities).filter((ent) => {
      if (!state.gmMode && ent.gmOnly) return false;
      return true;
    });
    const relPool = Object.values(camp.relationships).filter((rel) => {
      if (!state.gmMode && rel.secret) return false;
      return true;
    });
    // Build node list
    Object.values(camp.entities).forEach(ent => {
      if (!state.gmMode && ent.gmOnly) return;
      // Apply web type filter (GM only)
      if (state.gmMode && !webFilter[ent.type]) return;
      // Determine color and shape based on type
      const color = ent.type === 'pc' ? pcColor : ent.type === 'npc' ? npcColor : orgColor;
      const shape = ent.type === 'pc' ? 'circle' : ent.type === 'npc' ? 'rounded-rect' : 'hexagon';
      // Retrieve or assign a position in graph coordinates
      let pos = camp.positions[ent.id];
      if (!pos) {
        // Assign a random initial position within a 400x400 box centred at 0
        pos = { x: (Math.random() - 0.5) * 400, y: (Math.random() - 0.5) * 400 };
        camp.positions[ent.id] = { x: pos.x, y: pos.y };
      }
      const node = { id: ent.id, ent, x: pos.x, y: pos.y, color, shape };
      graphNodes.push(node);
      nodeMap[ent.id] = node;
    });
    // Build edge list
    Object.values(camp.relationships).forEach(rel => {
      const src = nodeMap[rel.source];
      const trg = nodeMap[rel.target];
      if (!src || !trg) return;
      if (!state.gmMode) {
        if (rel.secret) return;
      }
      // GM web filter: hide secrets if filtered
      if (state.gmMode && rel.secret && !webFilter.secrets) return;
      graphEdges.push({
        id: rel.id,
        source: src,
        target: trg,
        type: rel.type,
        directed: rel.directed,
        sourceKnows: rel.sourceKnows,
        targetKnows: rel.targetKnows,
        secret: rel.secret
      });
    });
    // Redraw the graph on the canvas
    drawGraph();
    // Refresh node select dropdown
    populateNodeSelect();
    populateGraphViewSelect();
    const emptyEl = document.getElementById('graph-empty');
    if (emptyEl) {
      const isEmpty = graphNodes.length === 0;
      emptyEl.classList.toggle('hidden', !isEmpty);
      if (isEmpty) {
        emptyEl.innerHTML = '';
        const icon = document.createElement('span');
        icon.className = 'material-icons';
        icon.textContent = 'hub';
        const text = document.createElement('p');
        text.textContent = state.gmMode
          ? 'No entities match current web filters. Add entities or relax filters above.'
          : 'No visible entities in the web yet.';
        emptyEl.appendChild(icon);
        emptyEl.appendChild(text);
        if (state.gmMode) {
          const actions = document.createElement('div');
          actions.className = 'graph-empty-actions';
          const filtersActive = !webFilter.pc || !webFilter.npc || !webFilter.org || !webFilter.secrets;
          if (filtersActive) {
            const clearFiltersBtn = document.createElement('button');
            clearFiltersBtn.className = 'toolbar-btn';
            clearFiltersBtn.textContent = 'Clear Web Filters';
            clearFiltersBtn.addEventListener('click', () => {
              webFilter.pc = true;
              webFilter.npc = true;
              webFilter.org = true;
              webFilter.secrets = true;
              syncWebFilterPills();
              updateGraph();
            });
            actions.appendChild(clearFiltersBtn);
          }
          const addPcBtn = document.createElement('button');
          addPcBtn.className = 'toolbar-btn';
          addPcBtn.textContent = 'Add PC';
          addPcBtn.addEventListener('click', () => {
            const btn = document.getElementById('add-pc-btn');
            if (btn) btn.click();
          });
          const addNpcBtn = document.createElement('button');
          addNpcBtn.className = 'toolbar-btn';
          addNpcBtn.textContent = 'Add NPC';
          addNpcBtn.addEventListener('click', () => {
            const btn = document.getElementById('add-npc-btn');
            if (btn) btn.click();
          });
          const addOrgBtn = document.createElement('button');
          addOrgBtn.className = 'toolbar-btn';
          addOrgBtn.textContent = 'Add Org';
          addOrgBtn.addEventListener('click', () => {
            const btn = document.getElementById('add-org-btn');
            if (btn) btn.click();
          });
          actions.appendChild(addPcBtn);
          actions.appendChild(addNpcBtn);
          actions.appendChild(addOrgBtn);
          emptyEl.appendChild(actions);
        }
      }
    }
    updateWebLegendMeta(camp, {
      totalEntities: entityPool.length,
      visibleEntities: graphNodes.length,
      totalRelationships: relPool.length,
      visibleRelationships: graphEdges.length
    });
  }

  /* -----------------------------------------------------------------------
   * Custom canvas-based graph utilities. These functions handle resizing,
   * drawing of nodes and edges, hit-testing, dragging, panning, zooming,
   * selection and layout. They replace the Cytoscape implementation used
   * in earlier versions. See setupGraph() for initialisation.
   */

  /** Resize the graph canvas to match its container. Also accounts for
   * device pixel ratio for crisp rendering. Re-draws the graph on
   * resize. */
  function resizeGraphCanvas() {
    const container = document.getElementById('cy-container');
    if (!graphCanvas || !container) return;
    const dpr = window.devicePixelRatio || 1;
    const width = container.clientWidth;
    const height = container.clientHeight;
    // Reset the transform when resizing to avoid compounding scale
    graphCtx.setTransform(1, 0, 0, 1, 0, 0);
    graphCanvas.width = width * dpr;
    graphCanvas.height = height * dpr;
    graphCanvas.style.width = width + 'px';
    graphCanvas.style.height = height + 'px';
    graphCtx.scale(dpr, dpr);
    drawGraph();
  }

  // Edge type → color mapping for the conspiracy web (light theme)
  const EDGE_COLORS_LIGHT = {
    'Ally':             '#4caf50',
    'Enemy':            '#e53935',
    'Rival':            '#ff8f00',
    'Romantic':         '#e91e8c',
    'Family':           '#8e44ad',
    'Blackmail':        '#c62828',
    'Targeting':        '#b71c1c',
    'Surveillance':     '#546e7a',
    'Informant':        '#0288d1',
    'Handler':          '#01579b',
    'Patron/Client':    '#6d4c41',
    'Employer/Employee':'#5d4037',
    'Member Of':        '#558b2f',
    'Owes / Debt':      '#f57f17',
    'Unknown/Unclear':  '#78909c'
  };

  // Dark-theme variants with brighter contrast for readability.
  const EDGE_COLORS_DARK = {
    'Ally':             '#66d47a',
    'Enemy':            '#ff6b6b',
    'Rival':            '#ffb74d',
    'Romantic':         '#ff66b2',
    'Family':           '#c792ea',
    'Blackmail':        '#ff5252',
    'Targeting':        '#ff8a80',
    'Surveillance':     '#90a4ae',
    'Informant':        '#4fc3f7',
    'Handler':          '#64b5f6',
    'Patron/Client':    '#bcaaa4',
    'Employer/Employee':'#a1887f',
    'Member Of':        '#aed581',
    'Owes / Debt':      '#ffd54f',
    'Unknown/Unclear':  '#b0bec5'
  };

  function edgeColor(type) {
    const isDark = document.body.classList.contains('dark');
    const palette = isDark ? EDGE_COLORS_DARK : EDGE_COLORS_LIGHT;
    return palette[type] || (isDark ? '#b0bec5' : '#888888');
  }

  function edgeDashForType(type) {
    if (type === 'Enemy' || type === 'Rival' || type === 'Targeting') {
      return [8 / graphState.scale, 4 / graphState.scale];
    }
    if (type === 'Surveillance') {
      return [2 / graphState.scale, 6 / graphState.scale];
    }
    return [];
  }

  /** Draw the conspiracy web on the canvas. */
  function drawGraph() {
    if (!graphCanvas || !graphCtx) return;
    const ctx = graphCtx;
    const dpr = window.devicePixelRatio || 1;
    const width = graphCanvas.width / dpr;
    const height = graphCanvas.height / dpr;
    ctx.clearRect(0, 0, width, height);
    ctx.save();
    ctx.translate(width / 2 + graphState.translateX, height / 2 + graphState.translateY);
    ctx.scale(graphState.scale, graphState.scale);

    const focusNodeId = getFocusedGraphNodeId();
    const focusNeighborIds = new Set();
    if (focusNodeId) {
      focusNeighborIds.add(focusNodeId);
      graphEdges.forEach(edge => {
        if (edge.source.id === focusNodeId) focusNeighborIds.add(edge.target.id);
        if (edge.target.id === focusNodeId) focusNeighborIds.add(edge.source.id);
      });
    }

    // Draw edges
    graphEdges.forEach(edge => {
      const x1 = edge.source.x, y1 = edge.source.y;
      const x2 = edge.target.x, y2 = edge.target.y;
      const isSelected = state.selectedRelId === edge.id;
      const color = edgeColor(edge.type);
      const focusEdge = !focusNodeId || edge.source.id === focusNodeId || edge.target.id === focusNodeId;
      ctx.strokeStyle = color;
      ctx.lineWidth = (isSelected ? 3 : 2) / graphState.scale;
      ctx.setLineDash(edgeDashForType(edge.type));
      if (focusNodeId) ctx.globalAlpha = isSelected ? 1 : (focusEdge ? 0.9 : 0.12);
      else ctx.globalAlpha = isSelected ? 1 : 0.75;
      if (edge.secret && !isSelected) ctx.globalAlpha *= 0.65;
      ctx.beginPath();
      ctx.moveTo(x1, y1);
      ctx.lineTo(x2, y2);
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.globalAlpha = 1;
      if (edge.directed) drawArrow(ctx, x1, y1, x2, y2, color, isSelected, focusNodeId && !focusEdge ? 0.18 : 1);
      // Edge label at midpoint
      const mx = (x1 + x2) / 2;
      const my = (y1 + y2) / 2;
      const edgeLabel = edge.secret ? `🔒 ${edge.type}` : edge.type;
      drawEdgeLabel(ctx, edgeLabel, mx, my, color, focusNodeId && !focusEdge ? 0.18 : 1);
    });

    // Draw nodes
    graphNodes.forEach(node => {
      const isSelected = graphSelectedNodeId === node.id;
      const isHovered = graphHoveredNodeId === node.id;
      if (focusNodeId) ctx.globalAlpha = focusNeighborIds.has(node.id) ? 1 : 0.2;
      else ctx.globalAlpha = 1;
      drawNodeShape(ctx, node.shape, node.x, node.y, node.color, isSelected, isHovered);
    });
    ctx.globalAlpha = 1;

    // Node labels
    ctx.font = `bold ${12 / graphState.scale}px 'Crimson Pro', Georgia, serif`;
    graphNodes.forEach(node => {
      if (focusNodeId) ctx.globalAlpha = focusNeighborIds.has(node.id) ? 1 : 0.25;
      else ctx.globalAlpha = 1;
      drawNodeLabel(ctx, entityLabel(node.ent), node.x, node.y);
    });
    ctx.globalAlpha = 1;
    ctx.restore();

    // Tooltip (drawn in screen space, outside transform)
    if (graphTooltip.visible) {
      drawTooltip(ctx, graphTooltip.text, graphTooltip.x, graphTooltip.y, width, height);
    }
  }

  function drawEdgeLabel(ctx, text, x, y, color, alpha = 1) {
    if (!text || graphState.scale < 0.5) return;
    ctx.save();
    ctx.globalAlpha = alpha;
    ctx.font = `${10 / graphState.scale}px 'Cinzel', serif`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    const metrics = ctx.measureText(text);
    const pad = 4 / graphState.scale;
    const bw = metrics.width + pad * 2;
    const bh = 13 / graphState.scale;
    // Background pill
    ctx.fillStyle = document.body.classList.contains('dark') ? 'rgba(20,16,12,0.82)' : 'rgba(245,240,232,0.88)';
    ctx.beginPath();
    const r = bh / 2;
    ctx.roundRect(x - bw/2, y - bh/2, bw, bh, [r]);
    ctx.fill();
    ctx.fillStyle = color;
    ctx.fillText(text, x, y);
    ctx.restore();
  }

  function drawTooltip(ctx, text, sx, sy, canvasW, canvasH) {
    ctx.save();
    ctx.font = '12px sans-serif';
    const pad = 8;
    const w = ctx.measureText(text).width + pad * 2;
    const h = 26;
    let tx = sx + 14;
    let ty = sy - h / 2;
    if (tx + w > canvasW - 8) tx = sx - w - 8;
    if (ty < 4) ty = 4;
    if (ty + h > canvasH - 4) ty = canvasH - h - 4;
    ctx.fillStyle = document.body.classList.contains('dark') ? '#2a2018' : '#fff8ee';
    ctx.strokeStyle = '#888';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.roundRect(tx, ty, w, h, [5]);
    ctx.fill();
    ctx.stroke();
    ctx.fillStyle = document.body.classList.contains('dark') ? '#e8ddd0' : '#1a1410';
    ctx.textBaseline = 'middle';
    ctx.fillText(text, tx + pad, ty + h / 2);
    ctx.restore();
  }

  /** Draw a node shape at graph coordinates. Supports circle (PC),
   * rounded rectangle (NPC) and hexagon (Org).
   * isSelected draws a bright selection ring, isHovered a subtle glow. */
  function drawNodeShape(ctx, shape, x, y, color, isSelected, isHovered) {
    const r = 18; // radius/half-size
    // Selection / hover ring
    if (isSelected || isHovered) {
      ctx.beginPath();
      const ringR = r + (isSelected ? 5 : 3);
      ctx.arc(x, y, ringR, 0, Math.PI * 2);
      ctx.fillStyle = isSelected ? 'rgba(255,255,255,0.35)' : 'rgba(255,255,255,0.15)';
      ctx.fill();
    }
    ctx.fillStyle = color;
    ctx.beginPath();
    if (shape === 'circle') {
      ctx.arc(x, y, r, 0, Math.PI * 2);
    } else if (shape === 'rounded-rect') {
      const radius = 6;
      const w = r * 2;
      const h = r * 2;
      const left = x - r;
      const top = y - r;
      ctx.moveTo(left + radius, top);
      ctx.lineTo(left + w - radius, top);
      ctx.quadraticCurveTo(left + w, top, left + w, top + radius);
      ctx.lineTo(left + w, top + h - radius);
      ctx.quadraticCurveTo(left + w, top + h, left + w - radius, top + h);
      ctx.lineTo(left + radius, top + h);
      ctx.quadraticCurveTo(left, top + h, left, top + h - radius);
      ctx.lineTo(left, top + radius);
      ctx.quadraticCurveTo(left, top, left + radius, top);
    } else if (shape === 'hexagon') {
      for (let i = 0; i < 6; i++) {
        const angle = Math.PI / 3 * i + Math.PI / 6;
        const px = x + r * Math.cos(angle);
        const py = y + r * Math.sin(angle);
        if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
      }
    }
    ctx.closePath();
    ctx.fill();
    // Outline
    ctx.strokeStyle = '#333';
    ctx.lineWidth = 1 / graphState.scale;
    ctx.stroke();
  }

  /** Draw a node label slightly offset below the node. */
  function drawNodeLabel(ctx, label, x, y) {
    if (!label) return;
    const isDark = document.body.classList.contains('dark');
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    const labelY = y + 22 / graphState.scale;
    if (isDark) {
      // White text with dark shadow for dark mode
      ctx.shadowColor = 'rgba(0,0,0,0.9)';
      ctx.shadowBlur = 4 / graphState.scale;
      ctx.shadowOffsetX = 0;
      ctx.shadowOffsetY = 1 / graphState.scale;
      ctx.fillStyle = '#ffffff';
    } else {
      // Black text with light shadow for light/parchment mode
      ctx.shadowColor = 'rgba(255,255,255,0.8)';
      ctx.shadowBlur = 3 / graphState.scale;
      ctx.shadowOffsetX = 0;
      ctx.shadowOffsetY = 0;
      ctx.fillStyle = '#1a1410';
    }
    ctx.fillText(label, x, labelY);
    ctx.shadowColor = 'transparent';
    ctx.shadowBlur = 0;
  }

  /** Draw an arrowhead at the target end of the edge. */
  function drawArrow(ctx, x1, y1, x2, y2, color, isSelected, alpha = 1) {
    const angle = Math.atan2(y2 - y1, x2 - x1);
    const arrowLength = 12 / graphState.scale;
    const arrowWidth = 6 / graphState.scale;
    ctx.save();
    ctx.translate(x2, y2);
    ctx.rotate(angle);
    ctx.beginPath();
    ctx.moveTo(0, 0);
    ctx.lineTo(-arrowLength, arrowWidth);
    ctx.lineTo(-arrowLength, -arrowWidth);
    ctx.closePath();
    ctx.fillStyle = isSelected ? '#fff' : (color || '#888');
    ctx.globalAlpha = (isSelected ? 1 : 0.75) * alpha;
    ctx.fill();
    ctx.globalAlpha = 1;
    ctx.restore();
  }

  /** Determine whether a given screen coordinate (px,py) hits a node.
   * Converts the point into graph space and checks distance to each node. */
  function getNodeAt(px, py) {
    const dpr = window.devicePixelRatio || 1;
    const width = graphCanvas.width / dpr;
    const height = graphCanvas.height / dpr;
    // Convert to graph coordinates
    const gx = (px - width / 2 - graphState.translateX) / graphState.scale;
    const gy = (py - height / 2 - graphState.translateY) / graphState.scale;
    const r = 18;
    for (let i = graphNodes.length - 1; i >= 0; i--) {
      const node = graphNodes[i];
      const dx = gx - node.x;
      const dy = gy - node.y;
      if (Math.sqrt(dx * dx + dy * dy) <= r) {
        return node;
      }
    }
    return null;
  }

  /** Determine whether a given screen coordinate (px,py) hits an edge.
   * Returns the first edge found within a small threshold. */
  function getEdgeAt(px, py) {
    const dpr = window.devicePixelRatio || 1;
    const width = graphCanvas.width / dpr;
    const height = graphCanvas.height / dpr;
    // Convert to graph coordinates
    const gx = (px - width / 2 - graphState.translateX) / graphState.scale;
    const gy = (py - height / 2 - graphState.translateY) / graphState.scale;
    const threshold = 6 / graphState.scale;
    for (let i = 0; i < graphEdges.length; i++) {
      const edge = graphEdges[i];
      const dist = pointToSegmentDistance(gx, gy, edge.source.x, edge.source.y, edge.target.x, edge.target.y);
      if (dist <= threshold) {
        return edge;
      }
    }
    return null;
  }

  /** Compute the distance from point (px,py) to a line segment (x1,y1)-(x2,y2). */
  function pointToSegmentDistance(px, py, x1, y1, x2, y2) {
    const dx = x2 - x1;
    const dy = y2 - y1;
    if (dx === 0 && dy === 0) {
      return Math.sqrt((px - x1) ** 2 + (py - y1) ** 2);
    }
    const t = ((px - x1) * dx + (py - y1) * dy) / (dx * dx + dy * dy);
    const tClamped = Math.max(0, Math.min(1, t));
    const lx = x1 + tClamped * dx;
    const ly = y1 + tClamped * dy;
    return Math.sqrt((px - lx) ** 2 + (py - ly) ** 2);
  }

  /** Mouse down event handler: start dragging a node or panning. */
  function onGraphMouseDown(e) {
    if (!graphCanvas) return;
    const rect = graphCanvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    const node = getNodeAt(x, y);
    graphState.movedDuringDrag = false;
    if (node) {
      graphState.draggingNodeId = node.id;
      graphState.dragOffsetX = node.x - ((x - rect.width / 2 - graphState.translateX) / graphState.scale);
      graphState.dragOffsetY = node.y - ((y - rect.height / 2 - graphState.translateY) / graphState.scale);
    } else {
      graphState.isPanning = true;
      graphState.lastX = e.clientX;
      graphState.lastY = e.clientY;
    }
  }

  /** Mouse move handler: drag node, pan view, or update hover tooltip. */
  function onGraphMouseMove(e) {
    if (!graphCanvas) return;
    const rect = graphCanvas.getBoundingClientRect();
    const mx = e.clientX - rect.left;
    const my = e.clientY - rect.top;
    if (graphState.draggingNodeId) {
      const node = graphNodes.find(n => n.id === graphState.draggingNodeId);
      if (!node) return;
      const x = (mx - rect.width / 2 - graphState.translateX) / graphState.scale;
      const y = (my - rect.height / 2 - graphState.translateY) / graphState.scale;
      node.x = x + graphState.dragOffsetX;
      node.y = y + graphState.dragOffsetY;
      currentCampaign().positions[node.id] = { x: node.x, y: node.y };
      graphState.movedDuringDrag = true;
      drawGraph();
    } else if (graphState.isPanning) {
      const dx = e.clientX - graphState.lastX;
      const dy = e.clientY - graphState.lastY;
      graphState.translateX += dx;
      graphState.translateY += dy;
      graphState.lastX = e.clientX;
      graphState.lastY = e.clientY;
      graphState.movedDuringDrag = true;
      drawGraph();
    } else {
      // Hover detection
      const hoveredNode = getNodeAt(mx, my);
      const newHoverId = hoveredNode ? hoveredNode.id : null;
      if (newHoverId !== graphHoveredNodeId) {
        graphHoveredNodeId = newHoverId;
        if (hoveredNode) {
          const ent = hoveredNode.ent;
          // Count relationships
          const relCount = Object.values(currentCampaign().relationships)
            .filter(r => r.source === ent.id || r.target === ent.id).length;
          const label = entityLabel(ent);
          const typeLabel = ent.type === 'pc' ? 'PC' : ent.type === 'npc' ? 'NPC' : 'Org';
          graphTooltip = {
            visible: true,
            text: label + ' (' + typeLabel + ') — ' + relCount + ' connection' + (relCount !== 1 ? 's' : ''),
            x: mx,
            y: my
          };
          graphCanvas.style.cursor = 'pointer';
        } else {
          graphTooltip = { visible: false, text: '', x: 0, y: 0 };
          graphCanvas.style.cursor = graphState.isPanning ? 'grabbing' : 'default';
        }
        drawGraph();
      } else if (hoveredNode) {
        // Update tooltip position
        graphTooltip.x = mx;
        graphTooltip.y = my;
        drawGraph();
      }
    }
  }

  /** Mouse up handler: end drag or pan. */
  function onGraphMouseUp(e) {
    if (!graphCanvas) return;
    graphState.draggingNodeId = null;
    graphState.isPanning = false;
  }

  /** Wheel handler: zoom in/out. */
  function onGraphWheel(e) {
    if (!graphCanvas) return;
    e.preventDefault();
    const rect = graphCanvas.getBoundingClientRect();
    const px = e.clientX - rect.left;
    const py = e.clientY - rect.top;
    const gx = px - rect.width / 2 - graphState.translateX;
    const gy = py - rect.height / 2 - graphState.translateY;
    const wheel = e.deltaY < 0 ? 1.1 : 0.9;
    const oldScale = graphState.scale;
    graphState.scale *= wheel;
    // Clamp zoom
    graphState.scale = Math.max(0.2, Math.min(4, graphState.scale));
    // Adjust translation to zoom around cursor
    graphState.translateX -= (gx / oldScale) * (graphState.scale - oldScale);
    graphState.translateY -= (gy / oldScale) * (graphState.scale - oldScale);
    drawGraph();
  }

  /** Click handler: if no drag occurred, select a node or edge. */
  function onGraphClick(e) {
    if (!graphCanvas) return;
    // If there was movement during drag, treat as drag not click
    if (graphState.movedDuringDrag) {
      graphState.movedDuringDrag = false;
      return;
    }
    const rect = graphCanvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    const node = getNodeAt(x, y);
    if (node) {
      selectEntity(node.id);
      return;
    }
    const edge = getEdgeAt(x, y);
    if (edge) {
      selectRelationship(edge.id);
    }
  }

  /** Right-click: show context menu for creating entities or connecting nodes. */
  function onGraphContextMenu(e) {
    e.preventDefault();
    if (!state.gmMode) return;
    const rect = graphCanvas.getBoundingClientRect();
    const mx = e.clientX - rect.left;
    const my = e.clientY - rect.top;
    // Remove existing context menu
    const existing = document.getElementById('graph-context-menu');
    if (existing) existing.remove();
    const menu = document.createElement('div');
    menu.id = 'graph-context-menu';
    menu.className = 'graph-context-menu';
    menu.style.left = (e.clientX) + 'px';
    menu.style.top = (e.clientY) + 'px';

    const clickedNode = getNodeAt(mx, my);
    const clickedEdge = clickedNode ? null : getEdgeAt(mx, my);

    if (clickedNode) {
      // Node context: connect from this node
      const connectItem = document.createElement('div');
      connectItem.className = 'ctx-item';
      connectItem.textContent = '🔗 Add relationship from ' + entityLabel(clickedNode.ent);
      connectItem.addEventListener('click', () => { menu.remove(); openRelModal(clickedNode.id); });
      menu.appendChild(connectItem);

      const selectItem = document.createElement('div');
      selectItem.className = 'ctx-item';
      selectItem.textContent = '📄 Open sheet';
      selectItem.addEventListener('click', () => { menu.remove(); selectEntity(clickedNode.id); document.querySelector('.tab-link[data-tab="sheets-view"]').click(); });
      menu.appendChild(selectItem);

      const deleteItem = document.createElement('div');
      deleteItem.className = 'ctx-item ctx-item-danger';
      deleteItem.textContent = '🗑️ Delete entity';
      deleteItem.addEventListener('click', async () => {
        menu.remove();
        await confirmAndDeleteEntity(clickedNode.id);
      });
      menu.appendChild(deleteItem);
    } else if (clickedEdge) {
      const openEdgeItem = document.createElement('div');
      openEdgeItem.className = 'ctx-item';
      openEdgeItem.textContent = '✏️ Edit relationship';
      openEdgeItem.addEventListener('click', () => {
        menu.remove();
        selectRelationship(clickedEdge.id);
      });
      menu.appendChild(openEdgeItem);

      const deleteEdgeItem = document.createElement('div');
      deleteEdgeItem.className = 'ctx-item ctx-item-danger';
      deleteEdgeItem.textContent = '🗑️ Delete relationship';
      deleteEdgeItem.addEventListener('click', async () => {
        menu.remove();
        await confirmAndDeleteRelationship(clickedEdge.id);
      });
      menu.appendChild(deleteEdgeItem);
    } else {
      // Canvas context: create entity at this position
      const dpr = window.devicePixelRatio || 1;
      const width = graphCanvas.width / dpr;
      const height = graphCanvas.height / dpr;
      const gx = (mx - width / 2 - graphState.translateX) / graphState.scale;
      const gy = (my - height / 2 - graphState.translateY) / graphState.scale;

      ['NPC','Org'].forEach(type => {
        const item = document.createElement('div');
        item.className = 'ctx-item';
        item.textContent = (type === 'NPC' ? '👤' : '🏛️') + ' Create ' + type + ' here';
        item.addEventListener('click', () => {
          menu.remove();
          const camp = currentCampaign();
          const ent = newEntity(camp, type.toLowerCase());
          ent.name = type === 'NPC' ? 'New NPC' : 'New Organisation';
          camp.positions[ent.id] = { x: gx, y: gy };
          appendLog('Created ' + type, ent.id);
          saveAndRefresh();
          selectEntity(ent.id);
        });
        menu.appendChild(item);
      });
    }

    const closeItem = document.createElement('div');
    closeItem.className = 'ctx-item ctx-item-cancel';
    closeItem.textContent = '✕ Cancel';
    closeItem.addEventListener('click', () => menu.remove());
    menu.appendChild(closeItem);

    document.body.appendChild(menu);
    // Dismiss on outside click
    setTimeout(() => {
      document.addEventListener('click', function dismiss() {
        menu.remove();
        document.removeEventListener('click', dismiss);
      });
    }, 0);
  }

  /** Double-click: jump to entity sheet. */
  function onGraphDblClick(e) {
    if (!graphCanvas) return;
    const rect = graphCanvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    const node = getNodeAt(x, y);
    if (node) {
      selectEntity(node.id);
      // Switch to sheets tab
      const sheetsTab = document.querySelector('.tab-link[data-tab="sheets-view"]');
      if (sheetsTab) sheetsTab.click();
    }
  }

  /** Center the view on a given node by adjusting translation. */
  function centerOnNode(nodeId) {
    const node = graphNodes.find(n => n.id === nodeId);
    if (!node || !graphCanvas) return;
    const dpr = window.devicePixelRatio || 1;
    const width = graphCanvas.width / dpr;
    const height = graphCanvas.height / dpr;
    graphState.translateX = -node.x * graphState.scale;
    graphState.translateY = -node.y * graphState.scale;
    drawGraph();
  }

  /** Run a simple force-directed layout to reposition nodes. This
   * algorithm applies repulsive and attractive forces for a set number
   * of iterations. Stores results in campaign.positions. */
  function forceLayout() {
    const iterations = 300;
    const k = 200; // spring length
    const repulsion = 40000;
    for (let i = 0; i < iterations; i++) {
      // Initialize forces
      const forces = {};
      graphNodes.forEach(n => forces[n.id] = { x: 0, y: 0 });
      // Repulsive forces between all nodes
      for (let a = 0; a < graphNodes.length; a++) {
        for (let b = a + 1; b < graphNodes.length; b++) {
          const n1 = graphNodes[a];
          const n2 = graphNodes[b];
          let dx = n1.x - n2.x;
          let dy = n1.y - n2.y;
          let dist = Math.sqrt(dx * dx + dy * dy) + 0.01;
          const force = repulsion / (dist * dist);
          const fx = force * dx / dist;
          const fy = force * dy / dist;
          forces[n1.id].x += fx;
          forces[n1.id].y += fy;
          forces[n2.id].x -= fx;
          forces[n2.id].y -= fy;
        }
      }
      // Attractive forces for edges
      graphEdges.forEach(edge => {
        const n1 = edge.source;
        const n2 = edge.target;
        let dx = n1.x - n2.x;
        let dy = n1.y - n2.y;
        let dist = Math.sqrt(dx * dx + dy * dy) + 0.01;
        const force = (dist * dist) / k;
        const fx = force * dx / dist;
        const fy = force * dy / dist;
        forces[n1.id].x -= fx;
        forces[n1.id].y -= fy;
        forces[n2.id].x += fx;
        forces[n2.id].y += fy;
      });
      // Apply forces with small time step
      const step = 0.002;
      graphNodes.forEach(node => {
        node.x += forces[node.id].x * step;
        node.y += forces[node.id].y * step;
      });
    }
    // Save positions back to campaign and redraw
    graphNodes.forEach(node => {
      currentCampaign().positions[node.id] = { x: node.x, y: node.y };
    });
    drawGraph();
  }

  /**
   * Render the inspector panel for editing a relationship. Shows fields
   * for type, directionality, awareness flags, secrecy and notes.
   */
  function renderInspectorForRelationship() {
    const inspector = document.getElementById('inspector');
    const content = document.getElementById('inspector-content');
    content.innerHTML = '';
    if (!state.selectedRelId) {
      inspector.classList.add('hidden');
      return;
    }
    const rel = currentCampaign().relationships[state.selectedRelId];
    if (!rel) {
      inspector.classList.add('hidden');
      return;
    }
    inspector.classList.remove('hidden');
    // Relationship editor
    const header = document.createElement('h3');
    const src = currentCampaign().entities[rel.source];
    const trg = currentCampaign().entities[rel.target];
    header.textContent = `Relationship: ${entityLabel(src)} → ${entityLabel(trg)}`;
    content.appendChild(header);
    // Type selector
    const typeField = document.createElement('div');
    typeField.className = 'inspector-field';
    const typeLabel = document.createElement('label');
    typeLabel.textContent = 'Type';
    const typeSelect = document.createElement('select');
    state.relTypes.forEach(t => {
      const opt = document.createElement('option');
      opt.value = t;
      opt.textContent = t;
      if (rel.type === t) opt.selected = true;
      typeSelect.appendChild(opt);
    });
    typeSelect.addEventListener('change', e => {
      pushRelationshipUndo(rel.id, 'Relationship type');
      rel.type = e.target.value;
      appendLog('Edited relationship type', rel.id);
      saveAndRefresh();
    });
    typeField.appendChild(typeLabel);
    typeField.appendChild(typeSelect);
    content.appendChild(typeField);
    // Directionality checkbox
    const dirField = document.createElement('div');
    dirField.className = 'inspector-field';
    const dirLabel = document.createElement('label');
    dirLabel.textContent = 'Directed';
    const dirChk = document.createElement('input');
    dirChk.type = 'checkbox';
    dirChk.checked = rel.directed;
    dirChk.addEventListener('change', e => {
      pushRelationshipUndo(rel.id, 'Relationship direction');
      rel.directed = e.target.checked;
      appendLog('Edited relationship direction', rel.id);
      saveAndRefresh();
    });
    dirField.appendChild(dirLabel);
    dirField.appendChild(dirChk);
    content.appendChild(dirField);
    // Awareness flags
    const awSec = document.createElement('div');
    awSec.className = 'inspector-field';
    awSec.innerHTML = '<label>Mutual Awareness</label>';
    const srcChk = document.createElement('input');
    srcChk.type = 'checkbox';
    srcChk.checked = rel.sourceKnows;
    srcChk.addEventListener('change', e => {
      pushRelationshipUndo(rel.id, 'Relationship awareness');
      rel.sourceKnows = e.target.checked;
      appendLog('Edited awareness', rel.id);
      saveAndRefresh();
    });
    const srcLabel = document.createElement('span');
    srcLabel.textContent = `${entityLabel(src)} knows`; 
    const trgChk = document.createElement('input');
    trgChk.type = 'checkbox';
    trgChk.checked = rel.targetKnows;
    trgChk.addEventListener('change', e => {
      pushRelationshipUndo(rel.id, 'Relationship awareness');
      rel.targetKnows = e.target.checked;
      appendLog('Edited awareness', rel.id);
      saveAndRefresh();
    });
    const trgLabel = document.createElement('span');
    trgLabel.textContent = `${entityLabel(trg)} knows`; 
    awSec.appendChild(srcChk);
    awSec.appendChild(srcLabel);
    awSec.appendChild(trgChk);
    awSec.appendChild(trgLabel);
    content.appendChild(awSec);
    // Secret toggle
    const secretField = document.createElement('div');
    secretField.className = 'inspector-field';
    const secretLabel = document.createElement('label');
    secretLabel.textContent = 'Secret (GM-only)';
    const secretChk = document.createElement('input');
    secretChk.type = 'checkbox';
    secretChk.checked = rel.secret;
    secretChk.addEventListener('change', e => {
      pushRelationshipUndo(rel.id, 'Relationship secrecy');
      rel.secret = e.target.checked;
      appendLog('Edited secrecy', rel.id);
      saveAndRefresh();
    });
    secretField.appendChild(secretLabel);
    secretField.appendChild(secretChk);
    content.appendChild(secretField);

    // Bond fallout level selector. Relevant when relationships represent bonds
    const falloutField = document.createElement('div');
    falloutField.className = 'inspector-field';
    const falloutLabel = document.createElement('label');
    falloutLabel.textContent = 'Bond Fallout';
    const falloutSelect = document.createElement('select');
    ['Minor', 'Moderate', 'Severe'].forEach(level => {
      const opt = document.createElement('option');
      opt.value = level;
      opt.textContent = level;
      if ((rel.falloutLevel || 'Minor') === level) opt.selected = true;
      falloutSelect.appendChild(opt);
    });
    falloutSelect.addEventListener('change', e => {
      pushRelationshipUndo(rel.id, 'Bond fallout');
      rel.falloutLevel = e.target.value;
      appendLog('Edited bond fallout', rel.id);
      saveAndRefresh();
    });
    falloutField.appendChild(falloutLabel);
    falloutField.appendChild(falloutSelect);
    content.appendChild(falloutField);

    // Bond/social consequence workflow helper.
    const sourcePc = src && src.type === 'pc' ? src : null;
    const targetPc = trg && trg.type === 'pc' ? trg : null;
    const candidatePcs = [sourcePc, targetPc].filter(Boolean);
    const consequenceField = document.createElement('div');
    consequenceField.className = 'inspector-field';
    consequenceField.classList.add('consequence-helper-field');
    const consequenceLabel = document.createElement('label');
    consequenceLabel.textContent = 'Consequence Helper';
    consequenceField.appendChild(consequenceLabel);

    const consequenceType = document.createElement('select');
    [
      { value: 'bond-fallout', label: 'Add Bond Fallout' },
      { value: 'social-stress', label: 'Apply Social Stress' },
      { value: 'followup-task', label: 'Create Follow-up Task' }
    ].forEach((optDef) => {
      const opt = document.createElement('option');
      opt.value = optDef.value;
      opt.textContent = optDef.label;
      consequenceType.appendChild(opt);
    });
    consequenceField.appendChild(consequenceType);

    const targetSelect = document.createElement('select');
    if (!candidatePcs.length) {
      const opt = document.createElement('option');
      opt.value = '';
      opt.textContent = 'No PC on this relationship';
      targetSelect.appendChild(opt);
      targetSelect.disabled = true;
    } else {
      candidatePcs.forEach((pcEnt) => {
        const opt = document.createElement('option');
        opt.value = pcEnt.id;
        opt.textContent = entityLabel(pcEnt);
        targetSelect.appendChild(opt);
      });
    }
    consequenceField.appendChild(targetSelect);

    const trackSelect = document.createElement('select');
    ['Mind', 'Shadow', 'Reputation', 'Silver'].forEach((track) => {
      const opt = document.createElement('option');
      opt.value = track;
      opt.textContent = track;
      trackSelect.appendChild(opt);
    });
    consequenceField.appendChild(trackSelect);

    const amountInput = document.createElement('input');
    amountInput.type = 'number';
    amountInput.min = '1';
    amountInput.max = '10';
    amountInput.value = '1';
    amountInput.placeholder = 'Amount';
    consequenceField.appendChild(amountInput);

    const detailInput = document.createElement('input');
    detailInput.type = 'text';
    detailInput.placeholder = 'Optional detail/note';
    consequenceField.appendChild(detailInput);

    const applyConsequenceBtn = document.createElement('button');
    applyConsequenceBtn.textContent = 'Apply Consequence';
    applyConsequenceBtn.className = 'toolbar-btn';
    applyConsequenceBtn.disabled = !candidatePcs.length;
    applyConsequenceBtn.addEventListener('click', () => {
      if (!targetSelect.value) return;
      const ok = applyRelationshipConsequence(rel, {
        kind: consequenceType.value,
        targetId: targetSelect.value,
        track: trackSelect.value,
        amount: amountInput.value,
        severity: falloutSelect.value,
        detail: detailInput.value
      });
      if (!ok) return;
      showToast('Relationship consequence applied.', 'info');
      saveAndRefresh();
    });
    consequenceField.appendChild(applyConsequenceBtn);

    const syncConsequenceFieldState = () => {
      const type = consequenceType.value;
      const needsTrack = type === 'bond-fallout' || type === 'social-stress';
      const needsAmount = type === 'social-stress';
      trackSelect.style.display = needsTrack ? '' : 'none';
      amountInput.style.display = needsAmount ? '' : 'none';
    };
    consequenceType.addEventListener('change', syncConsequenceFieldState);
    syncConsequenceFieldState();
    content.appendChild(consequenceField);

    // Notes
    const notesField = document.createElement('div');
    notesField.className = 'inspector-field';
    const notesLabel = document.createElement('label');
    notesLabel.textContent = 'Notes';
    const notesInput = document.createElement('textarea');
    notesInput.value = rel.notes || '';
    notesInput.addEventListener('input', e => {
      rel.notes = e.target.value;
      queueDeferredSave(`rel-notes:${rel.id}`);
    });
    notesInput.addEventListener('change', e => {
      pushRelationshipUndo(rel.id, 'Relationship notes');
      rel.notes = e.target.value;
      appendLog('Edited relationship notes', rel.id);
      saveWithoutRefresh();
    });
    notesField.appendChild(notesLabel);
    notesField.appendChild(notesInput);
    content.appendChild(notesField);
    // Actions: delete relationship
    const actions = document.createElement('div');
    actions.className = 'inspector-actions';
    const undoBtn = document.createElement('button');
    undoBtn.textContent = 'Undo Relationship Edit';
    const relUndoStack = (currentCampaign().relationshipUndo && currentCampaign().relationshipUndo[rel.id]) || [];
    undoBtn.disabled = !(Array.isArray(relUndoStack) && relUndoStack.length);
    if (Array.isArray(relUndoStack) && relUndoStack.length) {
      const last = relUndoStack[relUndoStack.length - 1];
      const when = last.time ? new Date(last.time).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '';
      undoBtn.title = `Undo: ${last.label || 'last relationship edit'}${when ? ` (${when})` : ''}`;
    } else {
      undoBtn.title = 'No relationship edits to undo.';
    }
    undoBtn.addEventListener('click', () => {
      undoRelationshipEdit(rel.id);
    });
    actions.appendChild(undoBtn);
    const redoBtn = document.createElement('button');
    redoBtn.textContent = 'Redo Relationship Edit';
    const relRedoStack = (currentCampaign().relationshipRedo && currentCampaign().relationshipRedo[rel.id]) || [];
    redoBtn.disabled = !(Array.isArray(relRedoStack) && relRedoStack.length);
    if (Array.isArray(relRedoStack) && relRedoStack.length) {
      const last = relRedoStack[relRedoStack.length - 1];
      const when = last.time ? new Date(last.time).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '';
      redoBtn.title = `Redo: ${last.label || 'last relationship edit'}${when ? ` (${when})` : ''}`;
    } else {
      redoBtn.title = 'No relationship edits to redo.';
    }
    redoBtn.addEventListener('click', () => {
      redoRelationshipEdit(rel.id);
    });
    actions.appendChild(redoBtn);
    const deleteBtn = document.createElement('button');
    deleteBtn.textContent = 'Delete Relationship';
    deleteBtn.addEventListener('click', async () => {
      await confirmAndDeleteRelationship(rel.id);
    });
    actions.appendChild(deleteBtn);
    content.appendChild(actions);
  }

  /**
   * Render the activity log. In player mode, hide log entries that
   * reference GM-only entities or secret relationships.
   */
  function buildOpenTasksPrepSection(pcs) {
    const tasksDiv = document.createElement('div');
    tasksDiv.className = 'prep-section';
    tasksDiv.innerHTML = '<strong>Open Tasks</strong>';
    let anyTasks = false;
    pcs.forEach(pc => {
      (pc.tasks || []).filter(t => t.status !== 'Done').forEach(t => {
        anyTasks = true;
        const tRow = document.createElement('div');
        tRow.className = 'prep-task-row';
        const nameBtn = document.createElement('button');
        nameBtn.className = 'prep-entity-link';
        nameBtn.textContent = entityLabel(pc) + ':';
        nameBtn.addEventListener('click', () => { selectEntity(pc.id); document.querySelector('.tab-link[data-tab="sheets-view"]').click(); });
        tRow.appendChild(nameBtn);
        const tdesc = document.createElement('span');
        tdesc.textContent = ' ' + (t.title || '(unnamed)') + ' [' + t.status + ']';
        tRow.appendChild(tdesc);
        tasksDiv.appendChild(tRow);
      });
    });
    if (!anyTasks) tasksDiv.innerHTML += '<p style="color:var(--spire-dim);font-size:0.82rem">No open tasks.</p>';
    return tasksDiv;
  }

  function deriveAutomationSuggestions(camp, stats = {}) {
    const session = camp.currentSession || 1;
    const recent = (camp.logs || [])
      .filter((e) => (e.session || 1) === session && e.type !== 'session')
      .slice(-80);
    const hasClockNamed = (pattern) => (camp.clocks || []).some((c) => pattern.test(String(c.name || '')));
    const suggestions = [];

    const consequenceActions = recent.filter((e) => /Applied social consequence stress|Applied bond fallout|Resolved fallout/.test(e.action || ''));
    const memberActions = recent.filter((e) => /Added member|Removed member/.test(e.action || ''));
    const relEdits = recent.filter((e) => /Edited relationship type|Edited relationship direction|Edited awareness|Edited secrecy|Added relationship|Deleted relationship/.test(e.action || ''));
    const falloutAdds = recent.filter((e) => /Added fallout|Fallout triggered/.test(e.action || ''));
    const taskSpikes = recent.filter((e) => /Added task|Edited task|Cleared done tasks/.test(e.action || ''));

    if ((stats.pendingFalloutCount || 0) >= 3 && !hasClockNamed(/fallout aftermath/i)) {
      suggestions.push({
        key: 'clock-fallout-aftermath',
        text: 'Multiple unresolved fallout entries detected. Track fallout cleanup with a dedicated clock.',
        cta: 'Create Fallout Clock',
        action: { type: 'create-clock', name: 'Fallout Aftermath', size: 6 }
      });
    }

    if ((stats.highStressPcCount || 0) >= 2) {
      suggestions.push({
        key: 'raise-ministry',
        text: 'Two or more PCs are in high-stress range. Escalate external pressure to keep pacing coherent.',
        cta: 'Increase Ministry',
        action: { type: 'raise-ministry', amount: 1 }
      });
    }

    if ((stats.openTasksCount || 0) >= 6 && !hasClockNamed(/operation pressure/i)) {
      suggestions.push({
        key: 'clock-operation-pressure',
        text: 'Large open task backlog suggests unresolved fronts. Use an operation pressure clock.',
        cta: 'Add Pressure Clock',
        action: { type: 'create-clock', name: 'Operation Pressure', size: 8 }
      });
    }

    if ((relEdits.length >= 3 || memberActions.length > 0) && !hasClockNamed(/faction backlash|faction response/i)) {
      suggestions.push({
        key: 'clock-faction-backlash',
        text: 'Recent relationship/member changes likely alter faction posture. Track response with a faction backlash clock.',
        cta: 'Add Faction Clock',
        action: { type: 'create-clock', name: 'Faction Backlash', size: 6 }
      });
    }

    if (consequenceActions.length > 0 || memberActions.length > 0 || relEdits.length > 0) {
      suggestions.push({
        key: 'review-web',
        text: 'Recent consequence and relationship activity suggests the conspiracy web needs updates.',
        cta: 'Open Web',
        action: { type: 'open-web' }
      });
    }

    if (falloutAdds.length >= 2 && !hasClockNamed(/medical|recovery|downtime/i)) {
      suggestions.push({
        key: 'clock-recovery',
        text: 'Fallout has stacked quickly this session. Add a recovery clock to pace downtime and treatment.',
        cta: 'Add Recovery Clock',
        action: { type: 'create-clock', name: 'Recovery Pressure', size: 4 }
      });
    }

    if (taskSpikes.length >= 5 && !hasClockNamed(/complication queue|consequence queue/i)) {
      suggestions.push({
        key: 'clock-complications',
        text: 'Task churn is high; convert unresolved items into a visible consequence queue.',
        cta: 'Add Complication Clock',
        action: { type: 'create-clock', name: 'Complication Queue', size: 6 }
      });
    }

    return suggestions.slice(0, 6);
  }

  function buildSessionSuggestionsPrepSection(camp, stats = {}) {
    const section = document.createElement('div');
    section.className = 'prep-section';
    section.innerHTML = '<strong>Suggested Updates</strong>';

    const suggestions = deriveAutomationSuggestions(camp, stats);

    if (!suggestions.length) {
      const p = document.createElement('p');
      p.className = 'text-muted';
      p.style.fontSize = '0.82rem';
      p.textContent = 'No immediate automation suggestions. Current state looks stable.';
      section.appendChild(p);
      return section;
    }

    function applySuggestionAction(action, options = {}) {
      const safeOnly = !!options.safeOnly;
      if (!action || !action.type) return { applied: false, skipped: true };
      if (safeOnly && action.type === 'raise-ministry') return { applied: false, skipped: true };
      if (action.type === 'create-clock') {
        const name = action.name || 'New Clock';
        const exists = (camp.clocks || []).some((c) => String(c.name || '').toLowerCase() === String(name).toLowerCase());
        if (exists) return { applied: false, skipped: true };
        camp.clocks.push({ id: generateId('clock'), name, current: 0, size: Number(action.size) || 6 });
        appendLog('Added clock', '');
        return { applied: true };
      }
      if (action.type === 'raise-ministry') {
        camp.ministryAttention = Math.min(10, (camp.ministryAttention || 0) + (Number(action.amount) || 1));
        appendLog('Adjusted ministry attention', '');
        return { applied: true };
      }
      if (action.type === 'open-web') {
        const webTab = document.querySelector('.tab-link[data-tab="web-view"]');
        if (webTab) webTab.click();
        return { applied: true };
      }
      return { applied: false, skipped: true };
    }

    const batchRow = document.createElement('div');
    batchRow.className = 'prep-task-row';
    batchRow.style.justifyContent = 'flex-end';
    batchRow.style.marginBottom = '6px';
    const safeBtn = document.createElement('button');
    safeBtn.className = 'toolbar-btn';
    safeBtn.textContent = 'Apply All Safe';
    safeBtn.title = 'Applies safe suggestions only (creates clocks and opens web; skips ministry escalation)';
    safeBtn.addEventListener('click', () => {
      let applied = 0;
      suggestions.forEach((s) => {
        const result = applySuggestionAction(s.action || {}, { safeOnly: true });
        if (result.applied) applied += 1;
      });
      saveCampaigns();
      renderSessionPrep();
      if (applied > 0) showToast(`Applied ${applied} safe suggestion${applied === 1 ? '' : 's'}.`, 'info');
      else showToast('No safe suggestions to apply.', 'info');
    });
    batchRow.appendChild(safeBtn);
    section.appendChild(batchRow);

    suggestions.forEach((s) => {
      const row = document.createElement('div');
      row.className = 'prep-task-row';
      row.style.alignItems = 'center';
      row.style.justifyContent = 'space-between';
      const txt = document.createElement('span');
      txt.textContent = s.text;
      const btn = document.createElement('button');
      btn.className = 'toolbar-btn';
      btn.textContent = s.cta;
      btn.addEventListener('click', () => {
        const result = applySuggestionAction(s.action || {}, { safeOnly: false });
        if (!result.applied) {
          showToast('Suggestion already satisfied.', 'info');
          return;
        }
        saveCampaigns();
        renderSessionPrep();
      });
      row.appendChild(txt);
      row.appendChild(btn);
      section.appendChild(row);
    });
    return section;
  }

  function generateScenePromptData() {
    const pools = getEffectiveScenePromptPools(currentCampaign());
    return {
      complication: randomFrom(pools.complications),
      factionReaction: randomFrom(pools.factionReactions),
      twist: randomFrom(pools.twists),
      time: new Date().toISOString()
    };
  }

  function renderSessionPrep() {
    const camp = currentCampaign();
    const prepEl = document.getElementById('session-prep-panel');
    if (!prepEl) return;
    if (!state.gmMode) { prepEl.style.display = 'none'; return; }
    prepEl.style.display = 'block';
    prepEl.innerHTML = '';
    if (!Array.isArray(camp.clocks)) camp.clocks = [];

    const h = document.createElement('div');
    h.className = 'session-prep-header';
    h.textContent = 'Session Prep';
    prepEl.appendChild(h);
    const tasksOnlyWrap = document.createElement('label');
    tasksOnlyWrap.className = 'sidebar-pin-only-row';
    tasksOnlyWrap.style.marginBottom = '6px';
    const tasksOnlyChk = document.createElement('input');
    tasksOnlyChk.type = 'checkbox';
    tasksOnlyChk.checked = !!camp.sessionPrepTasksOnly;
    tasksOnlyChk.addEventListener('change', () => {
      camp.sessionPrepTasksOnly = !!tasksOnlyChk.checked;
      saveCampaigns();
      renderSessionPrep();
    });
    const tasksOnlyTxt = document.createElement('span');
    tasksOnlyTxt.textContent = 'Open tasks only view';
    tasksOnlyWrap.appendChild(tasksOnlyChk);
    tasksOnlyWrap.appendChild(tasksOnlyTxt);
    prepEl.appendChild(tasksOnlyWrap);
    const tasksOnlyMode = !!camp.sessionPrepTasksOnly;

    // Ministry Attention tracker
    const minRow = document.createElement('div');
    minRow.className = 'ministry-row';
    const minLabel = document.createElement('span');
    minLabel.textContent = 'Ministry Attention: ';
    const minVal = document.createElement('strong');
    minVal.textContent = (camp.ministryAttention || 0) + '/10';
    minVal.style.color = camp.ministryAttention >= 7 ? 'var(--gm-accent-hi)' : camp.ministryAttention >= 4 ? 'var(--pl-accent-hi)' : 'var(--spire-text)';
    const minMinus = document.createElement('button');
    minMinus.textContent = '−'; minMinus.className = 'ministry-btn';
    minMinus.addEventListener('click', () => { camp.ministryAttention = Math.max(0, (camp.ministryAttention || 0) - 1); saveCampaigns(); renderSessionPrep(); });
    const minPlus = document.createElement('button');
    minPlus.textContent = '+'; minPlus.className = 'ministry-btn';
    minPlus.addEventListener('click', () => { camp.ministryAttention = Math.min(10, (camp.ministryAttention || 0) + 1); saveCampaigns(); renderSessionPrep(); });
    minRow.appendChild(minLabel);
    minRow.appendChild(minMinus);
    minRow.appendChild(minVal);
    minRow.appendChild(minPlus);
    // Ministry bar
    const minBar = document.createElement('div');
    minBar.className = 'ministry-bar';
    const minFill = document.createElement('div');
    minFill.className = 'ministry-fill';
    minFill.style.width = ((camp.ministryAttention || 0) * 10) + '%';
    minBar.appendChild(minFill);
    prepEl.appendChild(minRow);
    prepEl.appendChild(minBar);

    // Session control summary
    const summary = document.createElement('div');
    summary.className = 'prep-section';
    const pcs = Object.values(camp.entities).filter(e => e.type === 'pc');
    const pendingFalloutCount = pcs.reduce((sum, pc) => sum + (pc.fallout || []).filter(f => !f.resolved).length, 0);
    const highStressPcCount = pcs.filter(pc => {
      const tracks = ['blood','mind','silver','shadow','reputation'];
      return tracks.some((t) => {
        const filled = (pc.stressFilled && pc.stressFilled[t]) ? pc.stressFilled[t].length : 0;
        const slots = getTrackTotalSlots(pc, t);
        return filled >= Math.max(1, slots - 2);
      });
    }).length;
    const openTasksCount = pcs.reduce((sum, pc) => sum + (pc.tasks || []).filter(t => t.status !== 'Done').length, 0);
    const activeClockCount = camp.clocks.filter(c => (c.current || 0) < (c.size || 4)).length;
    summary.innerHTML = `<strong>Control Summary</strong>
      <div class="text-muted" style="font-size:0.84rem;margin-top:4px">
        Pending fallout: ${pendingFalloutCount} | High-stress PCs: ${highStressPcCount} | Open tasks: ${openTasksCount} | Active clocks: ${activeClockCount}
      </div>`;
    prepEl.appendChild(summary);
    prepEl.appendChild(buildSessionSuggestionsPrepSection(camp, {
      pendingFalloutCount,
      highStressPcCount,
      openTasksCount,
      activeClockCount
    }));
    if (!camp.lastScenePrompt) camp.lastScenePrompt = null;
    const sceneDiv = document.createElement('div');
    sceneDiv.className = 'prep-section';
    sceneDiv.innerHTML = '<strong>Scene Prompt</strong>';
    const sceneText = document.createElement('div');
    sceneText.style.fontSize = '0.84rem';
    sceneText.style.color = 'var(--spire-muted)';
    sceneText.style.marginTop = '6px';
    if (camp.lastScenePrompt) {
      sceneText.innerHTML = `Complication: ${camp.lastScenePrompt.complication}<br>Faction reaction: ${camp.lastScenePrompt.factionReaction}<br>Twist: ${camp.lastScenePrompt.twist}`;
    } else {
      sceneText.textContent = 'No prompt generated yet.';
    }
    const sceneBtnRow = document.createElement('div');
    sceneBtnRow.style.display = 'flex';
    sceneBtnRow.style.gap = '6px';
    sceneBtnRow.style.marginTop = '8px';
    const genSceneBtn = document.createElement('button');
    genSceneBtn.className = 'toolbar-btn';
    genSceneBtn.textContent = camp.lastScenePrompt ? 'Reroll Prompt' : 'Generate Prompt';
    genSceneBtn.addEventListener('click', () => {
      camp.lastScenePrompt = generateScenePromptData();
      appendLog('Generated scene prompt', '');
      saveCampaigns();
      renderSessionPrep();
    });
    sceneBtnRow.appendChild(genSceneBtn);
    if (camp.lastScenePrompt) {
      const copyBtn = document.createElement('button');
      copyBtn.className = 'toolbar-btn';
      copyBtn.textContent = 'Copy';
      copyBtn.addEventListener('click', async () => {
        const text = `Complication: ${camp.lastScenePrompt.complication}\nFaction reaction: ${camp.lastScenePrompt.factionReaction}\nTwist: ${camp.lastScenePrompt.twist}`;
        try {
          if (navigator.clipboard?.writeText) await navigator.clipboard.writeText(text);
          showToast('Scene prompt copied.', 'info');
        } catch (_) {
          showToast('Could not copy scene prompt.', 'warn');
        }
      });
      sceneBtnRow.appendChild(copyBtn);
    }
    sceneDiv.appendChild(sceneText);
    sceneDiv.appendChild(sceneBtnRow);
    prepEl.appendChild(sceneDiv);
    if (tasksOnlyMode) {
      prepEl.appendChild(buildOpenTasksPrepSection(pcs));
      return;
    }

    // Party stress overview
    const stressDiv = document.createElement('div');
    stressDiv.className = 'prep-section';
    stressDiv.innerHTML = '<strong>Party Stress</strong>';
    if (pcs.length === 0) {
      stressDiv.innerHTML += '<p style="color:var(--spire-dim);font-size:0.82rem">No PCs yet.</p>';
      const addPcBtn = document.createElement('button');
      addPcBtn.className = 'toolbar-btn';
      addPcBtn.style.marginTop = '6px';
      addPcBtn.textContent = 'Create First PC';
      addPcBtn.addEventListener('click', () => {
        const addBtn = document.getElementById('add-pc-btn');
        if (addBtn) addBtn.click();
      });
      stressDiv.appendChild(addPcBtn);
    } else {
      pcs.forEach(pc => {
        const pcRow = document.createElement('div');
        pcRow.className = 'prep-pc-row';
        const nameBtn = document.createElement('button');
        nameBtn.className = 'prep-entity-link';
        nameBtn.textContent = entityLabel(pc);
        nameBtn.addEventListener('click', () => { selectEntity(pc.id); document.querySelector('.tab-link[data-tab="sheets-view"]').click(); });
        pcRow.appendChild(nameBtn);
        const tracks = ['blood','mind','silver','shadow','reputation'];
        tracks.forEach(t => {
          const filled = (pc.stressFilled && pc.stressFilled[t]) ? pc.stressFilled[t].length : 0;
          const slots = getTrackTotalSlots(pc, t);
          if (filled === 0) return;
          const pill = document.createElement('span');
          pill.className = 'stress-pill stress-pill-' + t;
          pill.textContent = t.charAt(0).toUpperCase() + ': ' + filled + '/' + slots;
          if (filled >= slots) pill.classList.add('maxed');
          pcRow.appendChild(pill);
        });
        stressDiv.appendChild(pcRow);
      });
    }
    prepEl.appendChild(stressDiv);

    // High-risk PCs snapshot
    const riskDiv = document.createElement('div');
    riskDiv.className = 'prep-section';
    riskDiv.innerHTML = '<strong>High Risk PCs</strong>';
    let anyRisk = false;
    pcs.forEach((pc) => {
      const total = totalStressForFallout(pc);
      const tracks = ['blood','mind','silver','shadow','reputation'];
      const maxTrack = tracks.reduce((m, t) => {
        const v = (pc.stressFilled && pc.stressFilled[t]) ? pc.stressFilled[t].length : 0;
        return Math.max(m, v);
      }, 0);
      const activeFalloutCount = (pc.fallout || []).filter(f => !f.resolved).length;
      const risky = total >= 9 || maxTrack >= 8 || activeFalloutCount >= 2;
      if (!risky) return;
      anyRisk = true;
      const row = document.createElement('div');
      row.className = 'prep-task-row';
      const nameBtn = document.createElement('button');
      nameBtn.className = 'prep-entity-link';
      nameBtn.textContent = entityLabel(pc);
      nameBtn.addEventListener('click', () => {
        selectEntity(pc.id);
        document.querySelector('.tab-link[data-tab="sheets-view"]').click();
      });
      const info = document.createElement('span');
      info.textContent = ` — Total stress ${total}, peak track ${maxTrack}, active fallout ${activeFalloutCount}`;
      row.appendChild(nameBtn);
      row.appendChild(info);
      riskDiv.appendChild(row);
    });
    if (!anyRisk) riskDiv.innerHTML += '<p style="color:var(--spire-dim);font-size:0.82rem">No high-risk PCs right now.</p>';
    prepEl.appendChild(riskDiv);

    // Active fallout overview
    const falloutDiv = document.createElement('div');
    falloutDiv.className = 'prep-section';
    falloutDiv.innerHTML = '<strong>Active Fallout</strong>';
    let anyFallout = false;
    pcs.forEach(pc => {
      const active = (pc.fallout || []).filter(f => !f.resolved);
      active.forEach(f => {
        anyFallout = true;
        const fRow = document.createElement('div');
        fRow.className = 'prep-fallout-row';
        const nameBtn = document.createElement('button');
        nameBtn.className = 'prep-entity-link';
        nameBtn.textContent = entityLabel(pc);
        nameBtn.addEventListener('click', () => { selectEntity(pc.id); document.querySelector('.tab-link[data-tab="sheets-view"]').click(); });
        fRow.appendChild(nameBtn);
        const desc = document.createElement('span');
        desc.textContent = ' — ' + f.severity + ' ' + f.type + (f.name ? ': ' + f.name : '');
        fRow.appendChild(desc);
        falloutDiv.appendChild(fRow);
      });
    });
    if (!anyFallout) falloutDiv.innerHTML += '<p style="color:var(--spire-dim);font-size:0.82rem">No active fallout.</p>';
    prepEl.appendChild(falloutDiv);

    prepEl.appendChild(buildOpenTasksPrepSection(pcs));

    // Active clocks
    const clocksDiv = document.createElement('div');
    clocksDiv.className = 'prep-section';
    clocksDiv.innerHTML = '<strong>Active Clocks</strong>';
    const addClockBtn = document.createElement('button');
    addClockBtn.className = 'toolbar-btn';
    addClockBtn.textContent = '+ Add Clock';
    addClockBtn.style.marginTop = '6px';
    addClockBtn.addEventListener('click', async () => {
      const name = await askPrompt('Clock name:', '', { title: 'Add Clock', submitText: 'Add' });
      if (!name || !name.trim()) return;
      const sizeRaw = await askPrompt('Clock size (4, 6, 8, 10):', '6', { title: 'Clock Size', submitText: 'Create' });
      const parsed = parseInt(sizeRaw, 10);
      const size = [4, 6, 8, 10].includes(parsed) ? parsed : 6;
      camp.clocks.push({ id: generateId('clock'), name: name.trim(), current: 0, size });
      appendLog('Added clock', '');
      saveCampaigns();
      renderSessionPrep();
    });
    clocksDiv.appendChild(addClockBtn);

    if (!camp.clocks.length) {
      const p = document.createElement('p');
      p.style.color = 'var(--spire-dim)';
      p.style.fontSize = '0.82rem';
      p.textContent = 'No active clocks.';
      clocksDiv.appendChild(p);
    } else {
      camp.clocks.forEach((clock) => {
        const row = document.createElement('div');
        row.className = 'prep-task-row';
        row.style.marginTop = '6px';
        const label = document.createElement('span');
        const size = clock.size || 4;
        const progress = Math.min(size, Math.max(0, clock.current || 0));
        label.textContent = `${clock.name || '(unnamed)'} — ${progress}/${size}`;
        row.appendChild(label);

        const minus = document.createElement('button');
        minus.className = 'ministry-btn';
        minus.textContent = '−';
        minus.title = 'Decrease clock';
        minus.addEventListener('click', () => {
          clock.current = Math.max(0, (clock.current || 0) - 1);
          appendLog('Adjusted clock', '');
          saveCampaigns();
          renderSessionPrep();
        });

        const plus = document.createElement('button');
        plus.className = 'ministry-btn';
        plus.textContent = '+';
        plus.title = 'Increase clock';
        plus.addEventListener('click', () => {
          clock.current = Math.min(size, (clock.current || 0) + 1);
          appendLog('Adjusted clock', '');
          saveCampaigns();
          renderSessionPrep();
        });

        const del = document.createElement('button');
        del.className = 'ministry-btn';
        del.textContent = '×';
        del.title = 'Delete clock';
        del.addEventListener('click', async () => {
          const ok = await askConfirm(`Delete clock "${clock.name || '(unnamed)'}"?`, 'Delete Clock');
          if (!ok) return;
          camp.clocks = camp.clocks.filter(c => c.id !== clock.id);
          appendLog('Deleted clock', '');
          saveCampaigns();
          renderSessionPrep();
        });

        row.appendChild(minus);
        row.appendChild(plus);
        row.appendChild(del);
        clocksDiv.appendChild(row);
      });
    }
    prepEl.appendChild(clocksDiv);
  }

  function syncLogFilterControls(camp) {
    const sessionSelect = document.getElementById('log-session-filter');
    const actorSelect = document.getElementById('log-actor-filter');
    const typeSelect = document.getElementById('log-type-filter');
    const searchInput = document.getElementById('log-search-input');
    if (searchInput) searchInput.value = logFilterState.query || '';
    if (!sessionSelect) return;

    const sessions = Array.from(new Set((camp.logs || []).map(e => String(e.session || 1)))).sort((a, b) => Number(b) - Number(a));
    const options = ['all', ...sessions];
    const current = logFilterState.session || 'all';
    if (!options.includes(current)) logFilterState.session = 'all';

    sessionSelect.innerHTML = '';
    const allOpt = document.createElement('option');
    allOpt.value = 'all';
    allOpt.textContent = 'All sessions';
    sessionSelect.appendChild(allOpt);
    sessions.forEach((session) => {
      const opt = document.createElement('option');
      opt.value = session;
      opt.textContent = `Session ${session}`;
      sessionSelect.appendChild(opt);
    });
    sessionSelect.value = logFilterState.session;

    if (actorSelect) {
      const actors = Array.from(new Set((camp.logs || []).map((e) => String(e.actor || '').trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b));
      const opts = ['all', ...actors];
      if (!opts.includes(logFilterState.actor)) logFilterState.actor = 'all';
      actorSelect.innerHTML = '';
      const allActor = document.createElement('option');
      allActor.value = 'all';
      allActor.textContent = 'All actors';
      actorSelect.appendChild(allActor);
      actors.forEach((a) => {
        const opt = document.createElement('option');
        opt.value = a;
        opt.textContent = a;
        actorSelect.appendChild(opt);
      });
      actorSelect.value = logFilterState.actor || 'all';
    }
    if (typeSelect) {
      typeSelect.value = logFilterState.actionType || 'all';
    }
  }

  function renderLog() {
    const logList = document.getElementById('log-list');
    logList.innerHTML = '';
    const camp = currentCampaign();
    syncLogFilterControls(camp);
    const query = (logFilterState.query || '').trim().toLowerCase();
    const selectedSession = logFilterState.session || 'all';
    const selectedActor = logFilterState.actor || 'all';
    const selectedType = logFilterState.actionType || 'all';
    const filtered = camp.logs.slice().reverse().filter((entry) => {
      const ent = camp.entities[entry.target] || camp.relationships[entry.target];
      if (!state.gmMode) {
        if (ent && ent.gmOnly) return false;
        if (ent && ent.secret) return false;
      }
      if (selectedSession !== 'all' && String(entry.session || 1) !== selectedSession) return false;
      if (selectedActor !== 'all' && String(entry.actor || '') !== selectedActor) return false;
      if (selectedType !== 'all' && String(entry.type || 'action') !== selectedType) return false;
      const targetLabel = resolveLogTargetLabel(camp, entry);
      const actorLabel = entry.actor || '';
      const searchable = `${entry.action || ''} ${targetLabel} ${(entry.type || '')} ${actorLabel}`.toLowerCase();
      if (query && !searchable.includes(query)) return false;
      return true;
    });
    const pageSize = Math.max(10, parseInt(logFilterState.pageSize, 10) || 50);
    logFilterState.pageSize = pageSize;
    const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
    logFilterState.page = Math.min(totalPages, Math.max(1, parseInt(logFilterState.page, 10) || 1));
    const start = (logFilterState.page - 1) * pageSize;
    const pageEntries = filtered.slice(start, start + pageSize);

    pageEntries.forEach(entry => {
      const li = document.createElement('li');
      const targetLabel = resolveLogTargetLabel(camp, entry);
      if (entry.type === 'session') {
        li.className = 'log-session-divider';
        li.textContent = entry.action;
        logList.appendChild(li);
        return;
      }
      if (entry.type === 'note') {
        li.className = 'log-note';
      }
      const time = document.createElement('time');
      time.dateTime = entry.time;
      time.textContent = new Date(entry.time).toLocaleString();
      li.appendChild(time);
      if (entry.session) {
        const sess = document.createElement('span');
        sess.className = 'log-session-tag';
        sess.textContent = 'S' + entry.session;
        li.appendChild(sess);
      }
      if (entry.actor) {
        const actor = document.createElement('span');
        actor.className = 'log-actor-tag';
        actor.textContent = entry.actorRole === 'gm' ? `GM:${entry.actor}` : entry.actor;
        li.appendChild(actor);
      }
      const span = document.createElement('span');
      span.textContent = ' ' + entry.action + (targetLabel ? ' — ' + targetLabel : '');
      li.appendChild(span);
      const targetEnt = entry.target ? camp.entities[entry.target] : null;
      const targetRel = entry.target ? camp.relationships[entry.target] : null;
      if (targetEnt || targetRel) {
        const jumpBtn = document.createElement('button');
        jumpBtn.className = 'toolbar-btn log-jump-btn';
        jumpBtn.textContent = 'Jump';
        jumpBtn.addEventListener('click', (e) => e.stopPropagation());
        jumpBtn.addEventListener('click', () => {
          if (targetEnt) {
            selectEntity(targetEnt.id);
            const sheetsTab = document.querySelector('.tab-link[data-tab="sheets-view"]');
            if (sheetsTab) sheetsTab.click();
            return;
          }
          if (targetRel) {
            const webTab = document.querySelector('.tab-link[data-tab="web-view"]');
            if (webTab) webTab.click();
            selectRelationship(targetRel.id);
          }
        });
        li.appendChild(jumpBtn);
      }
      li.addEventListener('click', () => openLogEntryDetailModal(entry, targetLabel));
      logList.appendChild(li);
    });
    renderLogPagination(filtered.length, totalPages);
    // Show empty state
    const empty = document.getElementById('log-empty');
    if (empty) {
      const noEntries = logList.children.length === 0;
      empty.classList.toggle('hidden', !noEntries);
      if (noEntries) {
        empty.innerHTML = '';
        const icon = document.createElement('span');
        icon.className = 'material-icons';
        icon.textContent = 'history';
        const p = document.createElement('p');
        const hasFilter = !!query || selectedSession !== 'all' || selectedActor !== 'all' || selectedType !== 'all';
        p.textContent = hasFilter
          ? 'No log entries match these filters.'
          : 'No activity yet. Edit a sheet, web relation, or message to create history.';
        empty.appendChild(icon);
        empty.appendChild(p);
        if (hasFilter) {
          const clearBtn = document.createElement('button');
          clearBtn.className = 'toolbar-btn';
          clearBtn.textContent = 'Clear Filters';
          clearBtn.addEventListener('click', () => {
            logFilterState.query = '';
            logFilterState.session = 'all';
            logFilterState.actor = 'all';
            logFilterState.actionType = 'all';
            logFilterState.page = 1;
            renderLog();
          });
          empty.appendChild(clearBtn);
        }
      }
    }
  }

  function toCsvCell(value) {
    const v = value == null ? '' : String(value);
    if (/[",\n]/.test(v)) return `"${v.replace(/"/g, '""')}"`;
    return v;
  }

  function exportFilteredLogCsv() {
    const camp = currentCampaign();
    const query = (logFilterState.query || '').trim().toLowerCase();
    const selectedSession = logFilterState.session || 'all';
    const selectedActor = logFilterState.actor || 'all';
    const selectedType = logFilterState.actionType || 'all';
    const rows = (camp.logs || []).filter((entry) => {
      const ent = camp.entities[entry.target] || camp.relationships[entry.target];
      if (!state.gmMode) {
        if (ent && ent.gmOnly) return false;
        if (ent && ent.secret) return false;
      }
      if (selectedSession !== 'all' && String(entry.session || 1) !== selectedSession) return false;
      if (selectedActor !== 'all' && String(entry.actor || '') !== selectedActor) return false;
      if (selectedType !== 'all' && String(entry.type || 'action') !== selectedType) return false;
      const targetLabel = resolveLogTargetLabel(camp, entry);
      const actorLabel = entry.actor || '';
      const searchable = `${entry.action || ''} ${targetLabel} ${(entry.type || '')} ${actorLabel}`.toLowerCase();
      if (query && !searchable.includes(query)) return false;
      return true;
    });
    const header = ['time', 'session', 'type', 'actor', 'actorRole', 'action', 'targetId', 'targetLabel'];
    const lines = [header.join(',')];
    rows.forEach((entry) => {
      lines.push([
        toCsvCell(entry.time || ''),
        toCsvCell(entry.session || ''),
        toCsvCell(entry.type || ''),
        toCsvCell(entry.actor || ''),
        toCsvCell(entry.actorRole || ''),
        toCsvCell(entry.action || ''),
        toCsvCell(entry.target || ''),
        toCsvCell(resolveLogTargetLabel(camp, entry))
      ].join(','));
    });
    const csv = lines.join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${(camp.name || 'campaign').replace(/\s+/g, '_')}_log.csv`;
    link.click();
  }

  function openLogEntryDetailModal(entry, targetLabel = '') {
    if (!entry) return;
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    const content = document.getElementById('modal-content');
    const titleEl = document.getElementById('modal-title');
    if (titleEl) titleEl.textContent = 'Log Entry';
    content.innerHTML = '';

    const fields = [
      ['Time', new Date(entry.time).toLocaleString()],
      ['Session', entry.session || ''],
      ['Type', entry.type || 'action'],
      ['Actor', entry.actor || 'Unknown'],
      ['Actor Role', entry.actorRole || ''],
      ['Action', entry.action || ''],
      ['Target', targetLabel || entry.targetLabel || '']
    ];
    fields.forEach(([k, v]) => {
      const row = document.createElement('div');
      row.className = 'modal-field modal-field-inline';
      const key = document.createElement('strong');
      key.textContent = `${k}:`;
      key.style.minWidth = '110px';
      const val = document.createElement('span');
      val.textContent = String(v || '');
      row.appendChild(key);
      row.appendChild(val);
      content.appendChild(row);
    });

    const btnRow = document.createElement('div');
    btnRow.style.display = 'flex';
    btnRow.style.gap = '8px';
    btnRow.style.marginTop = '10px';
    const closeBtn = document.createElement('button');
    closeBtn.className = 'modal-submit';
    closeBtn.textContent = 'Close';
    closeBtn.addEventListener('click', () => closeModal());
    btnRow.appendChild(closeBtn);
    content.appendChild(btnRow);
    overlay.classList.remove('hidden');
    modal.classList.remove('hidden');
  }

  function renderLogPagination(totalEntries, totalPages) {
    const pager = document.getElementById('log-pagination');
    if (!pager) return;
    pager.innerHTML = '';
    if (totalEntries <= 0) {
      pager.classList.add('hidden');
      return;
    }
    pager.classList.remove('hidden');
    const page = Math.min(totalPages, Math.max(1, parseInt(logFilterState.page, 10) || 1));
    const pageSize = Math.max(10, parseInt(logFilterState.pageSize, 10) || 50);
    const start = (page - 1) * pageSize + 1;
    const end = Math.min(totalEntries, start + pageSize - 1);

    const summary = document.createElement('span');
    summary.className = 'log-page-summary';
    summary.textContent = `Showing ${start}-${end} of ${totalEntries}`;
    pager.appendChild(summary);

    const pageSizeSel = document.createElement('select');
    [25, 50, 100].forEach((size) => {
      const opt = document.createElement('option');
      opt.value = String(size);
      opt.textContent = `${size}/page`;
      if (pageSize === size) opt.selected = true;
      pageSizeSel.appendChild(opt);
    });
    pageSizeSel.title = 'Log page size';
    pageSizeSel.addEventListener('change', (e) => {
      logFilterState.pageSize = parseInt(e.target.value, 10) || 50;
      logFilterState.page = 1;
      renderLog();
    });
    pager.appendChild(pageSizeSel);

    const prevBtn = document.createElement('button');
    prevBtn.className = 'toolbar-btn';
    prevBtn.textContent = 'Prev';
    prevBtn.disabled = page <= 1;
    prevBtn.addEventListener('click', () => {
      logFilterState.page = Math.max(1, page - 1);
      renderLog();
    });
    pager.appendChild(prevBtn);

    const pageInfo = document.createElement('span');
    pageInfo.className = 'log-page-summary';
    pageInfo.textContent = `Page ${page}/${totalPages}`;
    pager.appendChild(pageInfo);

    const nextBtn = document.createElement('button');
    nextBtn.className = 'toolbar-btn';
    nextBtn.textContent = 'Next';
    nextBtn.disabled = page >= totalPages;
    nextBtn.addEventListener('click', () => {
      logFilterState.page = Math.min(totalPages, page + 1);
      renderLog();
    });
    pager.appendChild(nextBtn);
  }

  function generateSessionRecapText() {
    const camp = currentCampaign();
    const session = camp.currentSession || 1;
    const sessionStart = (camp.logs || [])
      .filter(e => e.type === 'session' && (e.session || 1) === session)
      .map(e => new Date(e.time).getTime())
      .sort((a, b) => b - a)[0];
    const sessionStartMs = Number.isFinite(sessionStart) ? sessionStart : 0;

    const actions = (camp.logs || [])
      .filter(e => (e.session || 1) === session && e.type !== 'session')
      .slice(-12)
      .map((e) => {
        const label = resolveLogTargetLabel(camp, e);
        const actor = e.actor ? ` by ${e.actor}` : '';
        return `- ${e.action}${label ? ` (${label})` : ''}${actor}`;
      });

    const pcs = Object.values(camp.entities || {}).filter(e => e.type === 'pc');
    const falloutLines = [];
    pcs.forEach((pc) => {
      const active = (pc.fallout || []).filter(f => !f.resolved);
      if (!active.length) return;
      falloutLines.push(`- ${entityLabel(pc)}: ${active.length} active fallout`);
    });

    const taskLines = [];
    pcs.forEach((pc) => {
      const openTasks = (pc.tasks || []).filter(t => t.status !== 'Done');
      if (!openTasks.length) return;
      taskLines.push(`- ${entityLabel(pc)}: ${openTasks.length} open task(s)`);
    });

    const msgCount = (camp.messages || [])
      .filter(m => new Date(m.time).getTime() >= sessionStartMs)
      .length;

    const lines = [];
    lines.push(`Session ${session} recap for "${camp.name}"`);
    lines.push('');
    lines.push(`Messages this session: ${msgCount}`);
    lines.push(`Ministry Attention: ${camp.ministryAttention || 0}/10`);
    lines.push('');
    lines.push('Notable actions:');
    lines.push(...(actions.length ? actions : ['- No notable actions logged.']));
    lines.push('');
    lines.push('Active fallout:');
    lines.push(...(falloutLines.length ? falloutLines : ['- None.']));
    lines.push('');
    lines.push('Open tasks:');
    lines.push(...(taskLines.length ? taskLines : ['- None.']));
    return lines.join('\n');
  }

  function openSessionRecapModal() {
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    const content = document.getElementById('modal-content');
    const titleEl = document.getElementById('modal-title');
    if (titleEl) titleEl.textContent = 'Session Recap';
    content.innerHTML = '';

    const recap = generateSessionRecapText();
    const textarea = document.createElement('textarea');
    textarea.value = recap;
    textarea.rows = 16;
    textarea.style.width = '100%';
    content.appendChild(textarea);

    const row = document.createElement('div');
    row.style.display = 'flex';
    row.style.gap = '8px';
    row.style.marginTop = '10px';
    const copyBtn = document.createElement('button');
    copyBtn.className = 'modal-submit';
    copyBtn.textContent = 'Copy Recap';
    const closeBtn = document.createElement('button');
    closeBtn.textContent = 'Close';
    row.appendChild(copyBtn);
    row.appendChild(closeBtn);
    content.appendChild(row);

    copyBtn.addEventListener('click', async () => {
      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(textarea.value);
        } else {
          textarea.select();
          document.execCommand('copy');
        }
        showToast('Recap copied.', 'info');
      } catch (_) {
        showToast('Could not copy recap.', 'warn');
      }
    });
    closeBtn.addEventListener('click', () => closeModal());

    overlay.classList.remove('hidden');
    modal.classList.remove('hidden');
  }

  function openShortcutHelpModal() {
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    const content = document.getElementById('modal-content');
    const titleEl = document.getElementById('modal-title');
    if (titleEl) titleEl.textContent = 'Keyboard Shortcuts';
    content.innerHTML = '';

    const list = document.createElement('div');
    list.className = 'modal-form';
    const shortcuts = [
      ['/', 'Focus entity search'],
      ['Ctrl/Cmd + S', 'Save campaign'],
      ['Ctrl/Cmd + Z', 'Undo destructive action'],
      ['Ctrl/Cmd + Y', 'Redo undone action'],
      ['Ctrl/Cmd + Shift + Z', 'Redo undone action'],
      ['↑ / ↓', 'Select previous/next visible entity'],
      ['← / →', 'Switch between top tabs'],
      ['Shift + T', 'Add a task to selected PC'],
      ['?', 'Open this shortcut help'],
      ['Esc', 'Close modal/inspector']
    ];
    shortcuts.forEach(([keys, action]) => {
      const row = document.createElement('div');
      row.className = 'modal-field modal-field-inline';
      const k = document.createElement('code');
      k.textContent = keys;
      k.style.minWidth = '140px';
      const txt = document.createElement('span');
      txt.textContent = action;
      row.appendChild(k);
      row.appendChild(txt);
      list.appendChild(row);
    });

    const closeBtn = document.createElement('button');
    closeBtn.className = 'modal-submit';
    closeBtn.textContent = 'Close';
    closeBtn.style.marginTop = '10px';
    closeBtn.addEventListener('click', () => closeModal());
    content.appendChild(list);
    content.appendChild(closeBtn);

    overlay.classList.remove('hidden');
    modal.classList.remove('hidden');
  }

  function messageTargetLabel(target) {
    if (target === 'party') return 'Party';
    if (target === 'gm') return 'GM';
    if (typeof target === 'string' && target.startsWith('user:')) return target.slice(5);
    return 'Party';
  }

  function deriveGmWhisperTargets(camp = currentCampaign(), users = state.users, currentUser = state.currentUser) {
    const out = [];
    const seen = new Set();
    const gmUsers = new Set(Array.isArray(camp && camp.gmUsers) ? camp.gmUsers : []);
    const memberUsers = Array.isArray(camp && camp.memberUsers) ? camp.memberUsers : [];
    memberUsers.forEach((name) => {
      if (!name || name === currentUser) return;
      if (gmUsers.has(name)) return;
      if (seen.has(name)) return;
      seen.add(name);
      out.push(name);
    });
    if (!out.length) {
      Object.keys(users || {})
        .filter((name) => name && name !== currentUser)
        .sort((a, b) => a.localeCompare(b))
        .forEach((name) => {
          if (seen.has(name)) return;
          seen.add(name);
          out.push(name);
        });
    }
    return out.sort((a, b) => a.localeCompare(b));
  }

  function canViewMessage(msg) {
    if (state.gmMode) return true;
    const me = state.currentUser || '';
    const target = msg.target || 'party';
    if (target === 'party') return true;
    if (target === 'gm') return msg.fromUser === me;
    if (target.startsWith('user:')) {
      const targetUser = target.slice(5);
      return targetUser === me || msg.fromUser === me;
    }
    return false;
  }

  function messageStyleClass(msg) {
    const target = msg.target || 'party';
    if (target === 'party') return 'msg-party';
    if (target === 'gm') return 'msg-whisper';
    if (target.startsWith('user:')) return 'msg-whisper';
    return 'msg-party';
  }

  function appendMessage(target, text) {
    const camp = currentCampaign();
    if (!camp.messages) camp.messages = [];
    const me = state.currentUser || (state.gmMode ? 'GM' : 'Player');
    const readBy = {};
    readBy[me] = true;
    camp.messages.push({
      id: generateId('msg'),
      time: new Date().toISOString(),
      fromUser: me,
      fromRole: state.gmMode ? 'gm' : 'player',
      target: target || 'party',
      text: text || '',
      readBy
    });
  }

  function populateMessageTargets() {
    const sel = document.getElementById('message-target-select');
    if (!sel) return;
    const prior = sel.value;
    sel.innerHTML = '';
    const addOpt = (value, label) => {
      const opt = document.createElement('option');
      opt.value = value;
      opt.textContent = label;
      sel.appendChild(opt);
    };
    addOpt('party', 'Party');
    if (!state.gmMode) {
      addOpt('gm', 'Whisper to GM');
    } else {
      const players = deriveGmWhisperTargets(currentCampaign(), state.users, state.currentUser);
      players.forEach((name) => addOpt('user:' + name, 'Whisper to ' + name));
    }
    sel.value = Array.from(sel.options).some(o => o.value === prior) ? prior : 'party';
  }

  function currentUserKey() {
    return state.currentUser || (state.gmMode ? 'GM' : 'Player');
  }

  function isMessageUnreadForCurrentUser(msg) {
    const me = currentUserKey();
    if (!msg || msg.fromUser === me) return false;
    if (!msg.readBy || typeof msg.readBy !== 'object') return true;
    return !msg.readBy[me];
  }

  function messageMatchesFilter(msg) {
    const mode = messageFilterState.mode || 'all';
    if (mode === 'party') return (msg.target || 'party') === 'party';
    if (mode === 'whispers') return (msg.target || 'party') !== 'party';
    if (mode === 'unread') return isMessageUnreadForCurrentUser(msg);
    return true;
  }

  function markVisibleMessagesRead(visibleMessages) {
    const activeTab = document.querySelector('.tab-link.active');
    if (!activeTab || activeTab.dataset.tab !== 'messages-view') return false;
    const me = currentUserKey();
    let changed = false;
    (visibleMessages || []).forEach((msg) => {
      if (msg.fromUser === me) return;
      if (!msg.readBy || typeof msg.readBy !== 'object') msg.readBy = {};
      if (!msg.readBy[me]) {
        msg.readBy[me] = true;
        changed = true;
      }
    });
    return changed;
  }

  function updateMessagesUnreadBadge() {
    const badge = document.getElementById('messages-unread-badge');
    if (!badge) return;
    const camp = currentCampaign();
    if (!camp || !Array.isArray(camp.messages)) {
      badge.classList.add('hidden');
      return;
    }
    const unread = camp.messages
      .filter(canViewMessage)
      .filter(isMessageUnreadForCurrentUser)
      .length;
    badge.textContent = String(unread);
    badge.classList.toggle('hidden', unread <= 0);
  }

  function renderMessages() {
    const list = document.getElementById('messages-list');
    if (!list) return;
    const camp = currentCampaign();
    if (!camp.messages) camp.messages = [];
    const clearBtn = document.getElementById('clear-messages-btn');
    if (clearBtn) clearBtn.style.display = state.gmMode ? '' : 'none';
    populateMessageTargets();
    const modeBadge = document.getElementById('messages-mode-badge');
    if (modeBadge) modeBadge.textContent = state.gmMode ? 'GM View' : 'Player View';
    const filterSel = document.getElementById('messages-filter-select');
    if (filterSel) filterSel.value = messageFilterState.mode || 'all';
    list.innerHTML = '';
    const visible = camp.messages.filter(canViewMessage).filter(messageMatchesFilter);
    const readChanged = markVisibleMessagesRead(visible);
    visible.slice().reverse().forEach((msg) => {
      const row = document.createElement('div');
      row.className = 'message-row ' + messageStyleClass(msg);
      if (isMessageUnreadForCurrentUser(msg)) row.classList.add('unread');
      const meta = document.createElement('div');
      meta.className = 'message-meta';
      const from = msg.fromUser || (msg.fromRole === 'gm' ? 'GM' : 'Player');
      const toLabel = messageTargetLabel(msg.target || 'party');
      meta.textContent = `${new Date(msg.time).toLocaleString()} • ${from} → ${toLabel}`;
      const body = document.createElement('div');
      body.className = 'message-body';
      body.textContent = msg.text || '';
      row.appendChild(meta);
      row.appendChild(body);
      list.appendChild(row);
    });
    if (!visible.length) {
      const empty = document.createElement('div');
      empty.className = 'messages-empty';
      const mode = messageFilterState.mode || 'all';
      const message = mode === 'unread'
        ? 'No unread messages in this view.'
        : mode === 'whispers'
          ? 'No whispers yet. Use target select to whisper.'
          : mode === 'party'
            ? 'No party messages yet.'
            : 'No messages yet. Start with a party message or whisper.';
      const text = document.createElement('div');
      text.textContent = message;
      empty.appendChild(text);
      if (mode !== 'unread') {
        const quickBtn = document.createElement('button');
        quickBtn.className = 'toolbar-btn';
        quickBtn.textContent = mode === 'whispers' ? 'Start Whisper' : 'Start Party Message';
        quickBtn.style.marginTop = '8px';
        quickBtn.addEventListener('click', () => {
          const target = document.getElementById('message-target-select');
          const input = document.getElementById('message-input');
          if (target) target.value = mode === 'whispers' ? 'gm' : 'party';
          if (input) input.focus();
        });
        empty.appendChild(quickBtn);
      }
      list.appendChild(empty);
    }
    if (readChanged) saveCampaigns();
    updateMessagesUnreadBadge();
    updateUndoButtonState();
  }

  const GM_NOTE_COLORS = ['#f6df8d', '#f7caa8', '#b8e3c7', '#b9d5f5', '#e7c6f2', '#f3b9bf'];

  function ensureGmNotesStore(camp = currentCampaign()) {
    if (!camp) return [];
    if (!Array.isArray(camp.gmNotes)) camp.gmNotes = [];
    camp.gmNotes.forEach((note) => {
      if (!note.id) note.id = generateId('gmnote');
      if (!note.title) note.title = '';
      if (!note.body) note.body = '';
      if (!note.scope) note.scope = 'session';
      if (note.scope !== 'global' && note.scope !== 'session') note.scope = 'session';
      if (note.scope === 'session') {
        const fallbackSession = camp.currentSession || 1;
        note.session = Number.isFinite(parseInt(note.session, 10)) ? parseInt(note.session, 10) : fallbackSession;
      } else {
        note.session = null;
      }
      if (note.pinned !== true) note.pinned = false;
      if (!note.color) note.color = randomFrom(GM_NOTE_COLORS);
      if (!note.updatedAt) note.updatedAt = new Date().toISOString();
      if (!note.createdAt) note.createdAt = note.updatedAt;
    });
    return camp.gmNotes;
  }

  function createGmNote(camp, scope = 'session', sessionValue) {
    const sessionNum = Number.isFinite(parseInt(sessionValue, 10))
      ? parseInt(sessionValue, 10)
      : (camp.currentSession || 1);
    return {
      id: generateId('gmnote'),
      title: '',
      body: '',
      scope: scope === 'global' ? 'global' : 'session',
      session: scope === 'global' ? null : sessionNum,
      pinned: false,
      color: randomFrom(GM_NOTE_COLORS),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
  }

  function touchGmNote(note) {
    note.updatedAt = new Date().toISOString();
  }

  function gmNotesBySession(camp) {
    const notes = ensureGmNotesStore(camp);
    const global = notes.filter((n) => n.scope === 'global');
    const sessionMap = new Map();
    const maxSession = Math.max(1, camp.currentSession || 1);
    for (let i = 1; i <= maxSession; i += 1) sessionMap.set(i, []);
    notes.forEach((n) => {
      if (n.scope !== 'session') return;
      const s = Number.isFinite(parseInt(n.session, 10)) ? parseInt(n.session, 10) : (camp.currentSession || 1);
      if (!sessionMap.has(s)) sessionMap.set(s, []);
      sessionMap.get(s).push(n);
    });
    const sortFn = (a, b) => {
      if (!!a.pinned !== !!b.pinned) return a.pinned ? -1 : 1;
      const at = new Date(a.updatedAt || a.createdAt || 0).getTime();
      const bt = new Date(b.updatedAt || b.createdAt || 0).getTime();
      return bt - at;
    };
    global.sort(sortFn);
    sessionMap.forEach((arr) => arr.sort(sortFn));
    return { global, sessionMap };
  }

  function createGmNoteCard(camp, note, options = {}) {
    const card = document.createElement('article');
    card.className = 'gm-note-card';
    card.style.setProperty('--note-color', note.color || randomFrom(GM_NOTE_COLORS));

    const titleInput = document.createElement('input');
    titleInput.type = 'text';
    titleInput.placeholder = 'Title';
    titleInput.value = note.title || '';
    titleInput.className = 'gm-note-title';
    titleInput.addEventListener('input', (e) => {
      note.title = e.target.value;
      touchGmNote(note);
      queueDeferredSave(`gm-note-title:${note.id}`);
    });
    titleInput.addEventListener('change', (e) => {
      note.title = e.target.value;
      touchGmNote(note);
      appendLog('Edited GM note', '');
      saveWithoutRefresh();
    });
    card.appendChild(titleInput);

    const bodyInput = document.createElement('textarea');
    bodyInput.placeholder = 'Write note...';
    bodyInput.value = note.body || '';
    bodyInput.className = 'gm-note-body';
    bodyInput.addEventListener('input', (e) => {
      note.body = e.target.value;
      touchGmNote(note);
      queueDeferredSave(`gm-note-body:${note.id}`);
    });
    bodyInput.addEventListener('change', (e) => {
      note.body = e.target.value;
      touchGmNote(note);
      appendLog('Edited GM note', '');
      saveWithoutRefresh();
    });
    card.appendChild(bodyInput);

    const footer = document.createElement('div');
    footer.className = 'gm-note-footer';

    const pinBtn = document.createElement('button');
    pinBtn.className = 'toolbar-btn';
    pinBtn.textContent = note.pinned ? 'Unpin' : 'Pin';
    pinBtn.addEventListener('click', () => {
      note.pinned = !note.pinned;
      touchGmNote(note);
      appendLog(note.pinned ? 'Pinned GM note' : 'Unpinned GM note', '');
      saveAndRefresh();
    });
    footer.appendChild(pinBtn);

    const scopeBtn = document.createElement('button');
    scopeBtn.className = 'toolbar-btn';
    if (note.scope === 'global') {
      scopeBtn.textContent = `Move to Session ${camp.currentSession || 1}`;
      scopeBtn.title = 'Move this note inside the current session folder';
    } else {
      scopeBtn.textContent = 'Pin Outside Session';
      scopeBtn.title = 'Move this note to the outside-session container';
    }
    scopeBtn.addEventListener('click', () => {
      if (note.scope === 'global') {
        note.scope = 'session';
        note.session = camp.currentSession || 1;
      } else {
        note.scope = 'global';
        note.session = null;
      }
      touchGmNote(note);
      appendLog('Moved GM note', '');
      saveAndRefresh();
    });
    footer.appendChild(scopeBtn);

    if (note.scope === 'session') {
      const sessSel = document.createElement('select');
      const maxSession = Math.max(1, camp.currentSession || 1);
      for (let s = 1; s <= maxSession; s += 1) {
        const opt = document.createElement('option');
        opt.value = String(s);
        opt.textContent = `S${s}`;
        if (s === Number(note.session || 1)) opt.selected = true;
        sessSel.appendChild(opt);
      }
      sessSel.addEventListener('change', (e) => {
        note.session = parseInt(e.target.value, 10) || (camp.currentSession || 1);
        touchGmNote(note);
        appendLog('Moved GM note', '');
        saveAndRefresh();
      });
      footer.appendChild(sessSel);
    }

    const delBtn = document.createElement('button');
    delBtn.className = 'row-remove-btn';
    delBtn.textContent = '×';
    delBtn.title = 'Delete note';
    delBtn.addEventListener('click', async () => {
      const ok = await askConfirm('Delete this GM note?', 'Delete GM Note');
      if (!ok) return;
      camp.gmNotes = ensureGmNotesStore(camp).filter((n) => n.id !== note.id);
      appendLog('Deleted GM note', '');
      saveAndRefresh();
    });
    footer.appendChild(delBtn);
    card.appendChild(footer);

    const updated = document.createElement('div');
    updated.className = 'gm-note-updated';
    updated.textContent = `Updated ${new Date(note.updatedAt || note.createdAt || Date.now()).toLocaleString()}`;
    card.appendChild(updated);
    return card;
  }

  function renderGMNotes() {
    const board = document.getElementById('gm-notes-board');
    if (!board) return;
    const page = document.getElementById('gm-notes-view');
    if (!state.gmMode) {
      if (page) page.classList.remove('active');
      board.innerHTML = '<div class="messages-empty">GM Notes are visible to GMs only.</div>';
      return;
    }
    const camp = currentCampaign();
    const notes = ensureGmNotesStore(camp);
    const grouped = gmNotesBySession(camp);
    board.innerHTML = '';

    const outsideSection = document.createElement('section');
    outsideSection.className = 'gm-notes-section';
    outsideSection.innerHTML = '<h3>Outside Session</h3>';
    const outsideGrid = document.createElement('div');
    outsideGrid.className = 'gm-note-grid';
    if (!grouped.global.length) {
      const empty = document.createElement('p');
      empty.className = 'text-muted';
      empty.textContent = 'No outside-session notes.';
      outsideGrid.appendChild(empty);
    } else {
      grouped.global.forEach((note) => outsideGrid.appendChild(createGmNoteCard(camp, note)));
    }
    outsideSection.appendChild(outsideGrid);
    board.appendChild(outsideSection);

    const sessions = Array.from(grouped.sessionMap.keys()).sort((a, b) => a - b);
    sessions.forEach((sessionNum) => {
      const folder = document.createElement('details');
      folder.className = 'gm-session-folder';
      if (sessionNum === (camp.currentSession || 1)) folder.open = true;
      const summary = document.createElement('summary');
      const count = grouped.sessionMap.get(sessionNum)?.length || 0;
      summary.textContent = `Session ${sessionNum} (${count})`;
      folder.appendChild(summary);

      const grid = document.createElement('div');
      grid.className = 'gm-note-grid';
      const notesInSession = grouped.sessionMap.get(sessionNum) || [];
      if (!notesInSession.length) {
        const empty = document.createElement('p');
        empty.className = 'text-muted';
        empty.textContent = 'No notes in this session folder.';
        grid.appendChild(empty);
      } else {
        notesInSession.forEach((note) => grid.appendChild(createGmNoteCard(camp, note)));
      }
      const addBtn = document.createElement('button');
      addBtn.className = 'toolbar-btn';
      addBtn.textContent = `Add Note to Session ${sessionNum}`;
      addBtn.addEventListener('click', () => {
        notes.push(createGmNote(camp, 'session', sessionNum));
        appendLog('Added GM note', '');
        saveAndRefresh();
      });
      grid.appendChild(addBtn);
      folder.appendChild(grid);
      board.appendChild(folder);
    });
    ensureAutoGrowTextareas(board);
  }

  /**
   * Perform a JSON export of the current campaign. If gmExport is
   * false, strip GM-only entities, secret edges and GM notes.
   */
  function exportCampaign(gmExport = false) {
    const camp = JSON.parse(JSON.stringify(currentCampaign()));
    if (!gmExport) {
      const removedEntities = new Set();
      // Remove gmOnly entities
      Object.keys(camp.entities).forEach(id => {
        if (camp.entities[id].gmOnly) {
          removedEntities.add(id);
          delete camp.entities[id];
        } else {
          // Remove gmNotes
          delete camp.entities[id].gmNotes;
        }
      });
      const removedRelationships = new Set();
      // Remove secret relationships and any pointing to removed entities
      Object.keys(camp.relationships).forEach(rid => {
        const rel = camp.relationships[rid];
        if (rel.secret) {
          removedRelationships.add(rid);
          delete camp.relationships[rid];
        } else if (!camp.entities[rel.source] || !camp.entities[rel.target]) {
          removedRelationships.add(rid);
          delete camp.relationships[rid];
        }
      });
      if (Array.isArray(camp.messages)) {
        camp.messages = camp.messages.filter(m => (m.target || 'party') === 'party');
      }
      camp.gmNotes = [];
      if (Array.isArray(camp.logs)) {
        camp.logs = camp.logs.filter((entry) => {
          if (!entry || !entry.target) return true;
          if (removedEntities.has(entry.target) || removedRelationships.has(entry.target)) return false;
          if (!camp.entities[entry.target] && !camp.relationships[entry.target]) return false;
          return true;
        });
      }
    }
    // Clean positions for removed nodes
    Object.keys(camp.positions).forEach(pid => {
      if (!camp.entities[pid]) delete camp.positions[pid];
    });
    return JSON.stringify(camp, null, 2);
  }

  function mergeCampaignData(importedCamp) {
    const camp = currentCampaign();
    if (!camp || !importedCamp) return { entities: 0, relationships: 0, messages: 0 };
    const entityIdMap = {};
    let addedEntities = 0;
    let addedRelationships = 0;
    let addedMessages = 0;
    const mergedEntityIds = [];

    Object.entries(importedCamp.entities || {}).forEach(([id, ent]) => {
      const sourceEnt = JSON.parse(JSON.stringify(ent));
      if (!camp.entities[id]) {
        camp.entities[id] = sourceEnt;
        entityIdMap[id] = id;
        mergedEntityIds.push(id);
        addedEntities += 1;
      } else {
        const newId = generateId(sourceEnt.type || 'ent');
        sourceEnt.id = newId;
        camp.entities[newId] = sourceEnt;
        entityIdMap[id] = newId;
        mergedEntityIds.push(newId);
        addedEntities += 1;
      }
    });

    // Remap internal entity references for merged entities when IDs were remapped.
    mergedEntityIds.forEach((eid) => {
      const ent = camp.entities[eid];
      if (!ent) return;
      if (Array.isArray(ent.members)) {
        ent.members = ent.members.map((mid) => entityIdMap[mid] || mid).filter((mid) => !!camp.entities[mid]);
      }
      if (typeof ent.affiliation === 'string' && ent.affiliation) {
        ent.affiliation = entityIdMap[ent.affiliation] || ent.affiliation;
        if (!camp.entities[ent.affiliation]) ent.affiliation = '';
      }
    });

    Object.values(importedCamp.relationships || {}).forEach((rel) => {
      const source = entityIdMap[rel.source] || rel.source;
      const target = entityIdMap[rel.target] || rel.target;
      if (!camp.entities[source] || !camp.entities[target]) return;
      const duplicate = Object.values(camp.relationships || {}).some((existing) => (
        existing.source === source &&
        existing.target === target &&
        existing.type === rel.type &&
        !!existing.secret === !!rel.secret
      ));
      if (duplicate) return;
      const rid = (rel.id && !camp.relationships[rel.id]) ? rel.id : generateId('rel');
      camp.relationships[rid] = Object.assign({}, JSON.parse(JSON.stringify(rel)), { id: rid, source, target });
      addedRelationships += 1;
    });

    Object.entries(importedCamp.positions || {}).forEach(([id, pos]) => {
      const mapped = entityIdMap[id] || id;
      if (mapped && camp.entities[mapped] && !camp.positions[mapped]) {
        camp.positions[mapped] = JSON.parse(JSON.stringify(pos));
      }
    });

    if (!Array.isArray(camp.messages)) camp.messages = [];
    (importedCamp.messages || []).forEach((msg) => {
      const cloned = JSON.parse(JSON.stringify(msg));
      cloned.id = (cloned.id && !camp.messages.some((m) => m.id === cloned.id)) ? cloned.id : generateId('msg');
      camp.messages.push(cloned);
      addedMessages += 1;
    });

    return { entities: addedEntities, relationships: addedRelationships, messages: addedMessages };
  }

  /**
   * Import a campaign JSON. Supports merge mode or import-as-new campaign.
   */
  async function importCampaign(jsonStr) {
    try {
      const data = JSON.parse(jsonStr);
      if (!data.entities || !data.relationships) {
        showToast('Invalid campaign file', 'warn');
        return;
      }
      const merge = await askConfirm(
        'Merge this file into the current campaign? Cancel imports as a separate campaign.',
        'Import Mode'
      );
      if (merge) {
        captureUndoSnapshot('Merge imported campaign');
        const merged = mergeCampaignData(data);
        saveAndRefresh();
        showToast(`Merged: ${merged.entities} entities, ${merged.relationships} relationships, ${merged.messages} messages.`, 'info');
        return;
      }
      const newId = generateId('camp');
      data.id = newId;
      data.name = data.name || 'Imported Campaign';
      data.owner = state.currentUser || null;
      data.gmUsers = state.currentUser ? [state.currentUser] : [];
      data.memberUsers = state.currentUser ? [state.currentUser] : [];
      data.inviteCode = '';
      data.sourceInviteCode = '';
      if (!Array.isArray(data.messages)) data.messages = [];
      if (!Array.isArray(data.gmNotes)) data.gmNotes = [];
      if (!Array.isArray(data.clocks)) data.clocks = [];
      if (!Array.isArray(data.undoStack)) data.undoStack = [];
      if (!Array.isArray(data.redoStack)) data.redoStack = [];
      state.campaigns[newId] = data;
      state.currentCampaignId = newId;
      saveCampaigns();
      initAfterLoad();
      showToast('Campaign imported successfully', 'info');
    } catch (e) {
      showToast('Failed to import campaign: ' + e.message, 'warn');
    }
  }

  /**
   * Export the currently visible graph to a PNG file. If full is true,
   * export the full graph bounds; otherwise export the current viewport.
   */
  function exportGraphPNG(full = false) {
    // Export the conspiracy web drawn on the custom canvas to a PNG. If
    // full is true, compute a bounding box around all nodes and draw
    // them to a new offscreen canvas before exporting. Otherwise
    // capture the current viewport. Hidden nodes and secret edges are
    // already excluded from graphNodes and graphEdges.
    if (!graphCanvas || !graphCtx) return;
    let canvasToExport;
    if (!full) {
      // Export current viewport by cloning the visible canvas
      canvasToExport = document.createElement('canvas');
      canvasToExport.width = graphCanvas.width;
      canvasToExport.height = graphCanvas.height;
      const ctx2 = canvasToExport.getContext('2d');
      ctx2.fillStyle = '#ffffff';
      ctx2.fillRect(0, 0, canvasToExport.width, canvasToExport.height);
      ctx2.drawImage(graphCanvas, 0, 0);
    } else {
      // Compute bounding box in graph coordinates
      let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
      graphNodes.forEach(n => {
        if (n.x < minX) minX = n.x;
        if (n.x > maxX) maxX = n.x;
        if (n.y < minY) minY = n.y;
        if (n.y > maxY) maxY = n.y;
      });
      if (!isFinite(minX)) return;
      // Add margin
      const margin = 40;
      const width = (maxX - minX) + margin * 2;
      const height = (maxY - minY) + margin * 2;
      canvasToExport = document.createElement('canvas');
      canvasToExport.width = width;
      canvasToExport.height = height;
      const ctx2 = canvasToExport.getContext('2d');
      ctx2.fillStyle = '#ffffff';
      ctx2.fillRect(0, 0, width, height);
      // Draw nodes and edges scaled to the bounding box
      graphEdges.forEach(edge => {
        const x1 = (edge.source.x - minX + margin);
        const y1 = (edge.source.y - minY + margin);
        const x2 = (edge.target.x - minX + margin);
        const y2 = (edge.target.y - minY + margin);
        ctx2.strokeStyle = '#888';
        ctx2.lineWidth = 2;
        // Set line dash based on type
        if (edge.type === 'Ally') ctx2.setLineDash([]);
        else if (edge.type === 'Enemy' || edge.type === 'Rival') ctx2.setLineDash([6, 4]);
        else if (edge.type === 'Surveillance') ctx2.setLineDash([2, 6]);
        else ctx2.setLineDash([]);
        ctx2.beginPath();
        ctx2.moveTo(x1, y1);
        ctx2.lineTo(x2, y2);
        ctx2.stroke();
        ctx2.setLineDash([]);
        // Draw arrowheads for directed relationships
        if (edge.directed) {
          drawArrow(ctx2, x1, y1, x2, y2, edge.sourceKnows, edge.targetKnows);
        }
      });
      // Draw nodes
      graphNodes.forEach(node => {
        const cx = (node.x - minX + margin);
        const cy = (node.y - minY + margin);
        drawNodeShape(ctx2, node.shape, cx, cy, node.color);
        drawNodeLabel(ctx2, node.ent ? entityLabel(node.ent) : '', cx, cy);
      });
    }
    // Create data URI and trigger download
    const dataUrl = canvasToExport.toDataURL('image/png');
    const link = document.createElement('a');
    link.href = dataUrl;
    link.download = `${currentCampaign().name.replace(/\s+/g, '_')}_conspiracy.png`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  function prepareSheetForPrint(sheetNode) {
    const clone = sheetNode.cloneNode(true);
    clone.classList.add('print-clean-sheet');

    clone.querySelectorAll('button, .section-help').forEach((el) => el.remove());

    clone.querySelectorAll('input[type="checkbox"]').forEach((chk) => {
      const mark = document.createElement('span');
      mark.className = 'print-check';
      mark.textContent = chk.checked ? '[x]' : '[ ]';
      chk.replaceWith(mark);
    });

    clone.querySelectorAll('input').forEach((inp) => {
      const type = (inp.getAttribute('type') || 'text').toLowerCase();
      if (['hidden', 'file', 'checkbox', 'radio', 'button', 'submit'].includes(type)) {
        inp.remove();
        return;
      }
      const span = document.createElement('span');
      span.className = 'print-field';
      span.textContent = inp.value || '—';
      inp.replaceWith(span);
    });

    clone.querySelectorAll('select').forEach((sel) => {
      const span = document.createElement('span');
      span.className = 'print-field';
      span.textContent = sel.options[sel.selectedIndex]?.textContent || sel.value || '—';
      sel.replaceWith(span);
    });

    clone.querySelectorAll('textarea').forEach((ta) => {
      const block = document.createElement('div');
      block.className = 'print-text-block';
      block.textContent = ta.value || '';
      ta.replaceWith(block);
    });

    return clone;
  }

  /**
   * Print a single entity sheet. Temporarily opens the sheet in a new
   * window and triggers print. Uses the same HTML structure but only
   * includes the selected entity.
   */
  function exportPCIdentityCard(entityId) {
    const camp = currentCampaign();
    const ent = camp && camp.entities ? camp.entities[entityId] : null;
    if (!ent || ent.type !== 'pc') return;
    const name = escapeHtml(entityLabel(ent));
    const cls = escapeHtml(ent.class || 'Unclassed');
    const pronouns = escapeHtml(ent.pronouns || '');
    const portrait = ent.image
      ? `<img src="${ent.image}" alt="Portrait" style="width:100%;height:220px;object-fit:cover;border-radius:8px;border:1px solid #665a49;">`
      : '<div style="width:100%;height:220px;border:1px dashed #665a49;border-radius:8px;display:flex;align-items:center;justify-content:center;color:#8d7f6c;">No portrait</div>';
    const domains = (ent.domains || []).map(d => d.name).filter(Boolean).slice(0, 4).join(', ');
    const skills = (ent.skills || []).map(s => s.name).filter(Boolean).slice(0, 6).join(', ');
    const html = `<!doctype html><html><head><meta charset="utf-8"><title>${name} ID Card</title></head>
      <body style="margin:0;padding:16px;font-family:Georgia,serif;background:#f4efe7;color:#231d18;">
        <div style="max-width:380px;border:2px solid #5f4f3f;border-radius:10px;padding:12px;background:#fffaf1;">
          <h1 style="margin:0 0 4px;font-size:1.35rem;">${name}</h1>
          <div style="margin:0 0 10px;color:#6a5847;font-size:0.95rem;">${cls}${pronouns ? ' • ' + pronouns : ''}</div>
          ${portrait}
          <div style="margin-top:10px;font-size:0.86rem;"><strong>Skills:</strong> ${escapeHtml(skills || '—')}</div>
          <div style="margin-top:6px;font-size:0.86rem;"><strong>Domains:</strong> ${escapeHtml(domains || '—')}</div>
        </div>
        <script>window.print();</script>
      </body></html>`;
    const w = window.open('', '_blank');
    if (!w) {
      showToast('Popup blocked. Allow popups to export card.', 'warn');
      return;
    }
    w.document.open();
    w.document.write(html);
    w.document.close();
  }

  async function printEntitySheet(entityId, options = {}) {
    const ent = currentCampaign().entities[entityId];
    if (!ent) return;
    const forcePlayerSafe = !!options.forcePlayerSafe;
    let includeGM = state.gmMode && !forcePlayerSafe;
    if (state.gmMode && !forcePlayerSafe) {
      includeGM = await askConfirm(
        'Include GM-only notes/details in printout? Cancel prints a player-safe copy.',
        'Print Mode'
      );
    }
    const prevMode = state.gmMode;
    const prevSelected = state.selectedEntityId;
    state.gmMode = !!includeGM;
    const win = window.open('', '_blank');
    const doc = win.document;
    doc.write('<html><head><title>Print</title>');
    doc.write('<link rel="stylesheet" href="styles.css">');
    doc.write('</head><body>');
    // Render the sheet for printing
    const div = document.createElement('div');
    div.className = 'sheet-container print-clean-sheet';
    if (ent.type === 'pc') {
      renderPCSheet(ent);
      const prepared = prepareSheetForPrint(document.getElementById('sheet-container'));
      div.innerHTML = prepared.innerHTML;
    } else if (ent.type === 'npc') {
      renderNPCSheet(ent);
      const prepared = prepareSheetForPrint(document.getElementById('sheet-container'));
      div.innerHTML = prepared.innerHTML;
    } else if (ent.type === 'org') {
      renderOrgSheet(ent);
      const prepared = prepareSheetForPrint(document.getElementById('sheet-container'));
      div.innerHTML = prepared.innerHTML;
    }
    doc.body.innerHTML = div.outerHTML;
    doc.body.classList.add('print-sheet-view');
    doc.close();
    state.gmMode = prevMode;
    state.selectedEntityId = prevSelected;
    applyModeClasses();
    renderSheetView();
    setTimeout(() => {
      win.print();
      win.close();
    }, 100);
  }

  /**
   * Open a modal for players to submit an NPC proposal for GM approval.
   * The NPC is created immediately with pendingApproval: true and is
   * visible to both the player and GM with a "Pending" badge.
   */
  // -----------------------------------------------------------------------
  //  DICE ROLLER
  // -----------------------------------------------------------------------
  function openDiceRollerForSkill(skillName, mastered, pcContext = null) {
    const skills = pcContext && Array.isArray(pcContext.skills)
      ? pcContext.skills.map((s) => s.name).filter(Boolean)
      : [];
    const domains = pcContext && Array.isArray(pcContext.domains)
      ? pcContext.domains.map((d) => d.name).filter(Boolean)
      : [];
    const equipmentTags = pcContext && Array.isArray(pcContext.inventory)
      ? Array.from(new Set(pcContext.inventory.flatMap((i) => Array.isArray(i.tags) ? i.tags : []).filter(Boolean)))
      : [];
    openDiceRoller(skillName, mastered ? 2 : 1, {
      allowDomainBonus: true,
      allowTagHooks: true,
      composer: {
        skill: skillName || '',
        mastered: !!mastered,
        skills,
        domains,
        equipmentTags
      },
      logTargetId: pcContext ? pcContext.id : ''
    });
  }

  function openDiceRollerForResistance(resistanceName, resistanceValue, pcContext = null) {
    const bonus = Number.isFinite(resistanceValue) ? resistanceValue : 0;
    openDiceRoller(`${resistanceName} Resistance`, 1, {
      dice: [{ sides: 10, label: 'D10', color: '#9c4221' }],
      initialDifficulty: 0,
      lockDifficulty: true,
      modifierLabel: 'Pool modifier:',
      helperText: `${resistanceName} +${bonus} grants free stress slots, not a roll modifier.`,
      logTargetId: pcContext ? pcContext.id : ''
    });
  }

  function openDiceRoller(contextLabel, poolBonus, options = {}) {
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    const content = document.getElementById('modal-content');
    const titleEl = document.getElementById('modal-title');
    if (titleEl) titleEl.textContent = contextLabel ? 'Rolling: ' + contextLabel : 'Dice Roller';
    content.innerHTML = '';
    overlay.classList.remove('hidden');
    modal.classList.remove('hidden');

    const wrap = document.createElement('div');
    wrap.className = 'dice-roller';

    // Dice buttons
    const dice = options.dice || [
      { sides: 3, label: 'D3', color: '#7c4dab' },
      { sides: 6, label: 'D6', color: '#2b6cb0' },
      { sides: 8, label: 'D8', color: '#276749' },
      { sides: 10, label: 'D10', color: '#9c4221' }
    ];

    const btnRow = document.createElement('div');
    btnRow.className = 'dice-btn-row';
    dice.forEach(d => {
      const btn = document.createElement('button');
      btn.className = 'dice-die-btn';
      btn.textContent = d.label;
      btn.style.setProperty('--die-color', d.color);
      btn.addEventListener('click', () => rollDie(d.sides, d.label, resultEl));
      btnRow.appendChild(btn);
    });
    wrap.appendChild(btnRow);

    // Difficulty modifier
    const modRow = document.createElement('div');
    modRow.className = 'dice-mod-row';
    const modLabel = document.createElement('label');
    modLabel.textContent = options.modifierLabel || 'Difficulty modifier:';
    const modSelect = document.createElement('select');
    modSelect.id = 'dice-difficulty';
    const initialDifficulty = Number.isInteger(options.initialDifficulty) ? options.initialDifficulty : 0;
    if (options.lockDifficulty) {
      const opt = document.createElement('option');
      opt.value = String(initialDifficulty);
      opt.textContent = `${initialDifficulty >= 0 ? '+' : ''}${initialDifficulty} (locked)`;
      modSelect.appendChild(opt);
      modSelect.disabled = true;
    } else {
      [
        { value: '0', label: 'Standard (no mod)' },
        { value: '-1', label: 'Difficult (−1 pool)' },
        { value: '-2', label: 'Very Difficult (−2 pool)' },
        { value: '1', label: 'Assisted (+1 pool)' },
        { value: '2', label: 'Two Assists (+2 pool)' }
      ].forEach(o => {
        const opt = document.createElement('option');
        opt.value = o.value;
        opt.textContent = o.label;
        modSelect.appendChild(opt);
      });
      if (!['-2','-1','0','1','2'].includes(String(initialDifficulty))) {
        const customOpt = document.createElement('option');
        customOpt.value = String(initialDifficulty);
        customOpt.textContent = `Custom (${initialDifficulty >= 0 ? '+' : ''}${initialDifficulty} pool)`;
        modSelect.appendChild(customOpt);
      }
      modSelect.value = String(initialDifficulty);
    }
    modRow.appendChild(modLabel);
    modRow.appendChild(modSelect);
    wrap.appendChild(modRow);
    let composerMasteredChk = null;
    let composerDomainChk = null;
    let composerDomainSel = null;
    let composerTagImpact = null;
    if (options.composer) {
      const comp = options.composer;
      const box = document.createElement('div');
      box.className = 'dice-composer';

      const title = document.createElement('div');
      title.className = 'dice-hist-title';
      title.textContent = 'Roll Composer';
      box.appendChild(title);

      const skillRow = document.createElement('div');
      skillRow.className = 'dice-mod-row';
      const skillLbl = document.createElement('label');
      skillLbl.textContent = 'Skill:';
      const skillVal = document.createElement('select');
      const skillOptions = (comp.skills && comp.skills.length) ? comp.skills : [comp.skill || '(custom)'];
      skillOptions.forEach((name) => {
        const opt = document.createElement('option');
        opt.value = name;
        opt.textContent = name;
        if (name === comp.skill) opt.selected = true;
        skillVal.appendChild(opt);
      });
      skillVal.disabled = true;
      skillRow.appendChild(skillLbl);
      skillRow.appendChild(skillVal);
      box.appendChild(skillRow);

      const masteryRow = document.createElement('div');
      masteryRow.className = 'dice-domain-row';
      const mLbl = document.createElement('label');
      mLbl.className = 'dice-domain-label';
      composerMasteredChk = document.createElement('input');
      composerMasteredChk.type = 'checkbox';
      composerMasteredChk.checked = !!comp.mastered;
      mLbl.appendChild(composerMasteredChk);
      mLbl.appendChild(document.createTextNode(' Mastery applies (+1 die)'));
      masteryRow.appendChild(mLbl);
      box.appendChild(masteryRow);

      const domainRow = document.createElement('div');
      domainRow.className = 'dice-domain-row';
      const dLbl = document.createElement('label');
      dLbl.className = 'dice-domain-label';
      composerDomainChk = document.createElement('input');
      composerDomainChk.type = 'checkbox';
      dLbl.appendChild(composerDomainChk);
      dLbl.appendChild(document.createTextNode(' Domain applies (+1 die)'));
      domainRow.appendChild(dLbl);
      if (Array.isArray(comp.domains) && comp.domains.length) {
        composerDomainSel = document.createElement('select');
        comp.domains.forEach((name) => {
          const opt = document.createElement('option');
          opt.value = name;
          opt.textContent = name;
          composerDomainSel.appendChild(opt);
        });
        composerDomainSel.disabled = true;
        composerDomainChk.addEventListener('change', () => {
          composerDomainSel.disabled = !composerDomainChk.checked;
        });
        domainRow.appendChild(composerDomainSel);
      }
      box.appendChild(domainRow);

      const tagRow = document.createElement('div');
      tagRow.className = 'dice-mod-row';
      const tagLbl = document.createElement('label');
      tagLbl.textContent = 'Equipment impact:';
      composerTagImpact = document.createElement('select');
      [
        { value: '0', label: 'No impact (0)' },
        { value: '1', label: 'Helpful tags (+1 die)' },
        { value: '-1', label: 'Hindering tags (-1 die)' }
      ].forEach((optData) => {
        const opt = document.createElement('option');
        opt.value = optData.value;
        opt.textContent = optData.label;
        composerTagImpact.appendChild(opt);
      });
      tagRow.appendChild(tagLbl);
      tagRow.appendChild(composerTagImpact);
      box.appendChild(tagRow);

      if (Array.isArray(comp.equipmentTags) && comp.equipmentTags.length) {
        const tagHint = document.createElement('div');
        tagHint.className = 'text-muted';
        tagHint.style.fontSize = '0.78rem';
        tagHint.textContent = 'Available tags: ' + comp.equipmentTags.join(', ');
        box.appendChild(tagHint);
      }
      wrap.appendChild(box);
    }
    // Mastered / base pool info
    if (poolBonus && poolBonus > 1) {
      const mastNote = document.createElement('div');
      mastNote.style.fontSize = '0.8rem';
      mastNote.style.color = 'var(--pl-accent-hi)';
      mastNote.style.marginBottom = '6px';
      mastNote.textContent = '★ Mastered — rolling extra die (keep best)';
      wrap.appendChild(mastNote);
    }
    if (options.helperText) {
      const helper = document.createElement('div');
      helper.style.fontSize = '0.8rem';
      helper.style.color = 'var(--spire-muted)';
      helper.style.marginTop = '-6px';
      helper.textContent = options.helperText;
      wrap.appendChild(helper);
    }
    let domainBonusChk = null;
    if (options.allowDomainBonus) {
      const domainRow = document.createElement('div');
      domainRow.className = 'dice-domain-row';
      const lbl = document.createElement('label');
      lbl.className = 'dice-domain-label';
      domainBonusChk = document.createElement('input');
      domainBonusChk.type = 'checkbox';
      lbl.appendChild(domainBonusChk);
      lbl.appendChild(document.createTextNode(' Add Domain bonus (+1 die)'));
      domainRow.appendChild(lbl);
      wrap.appendChild(domainRow);
    }
    let equipmentTagSelect = null;
    if (options.allowTagHooks) {
      const tagRow = document.createElement('div');
      tagRow.className = 'dice-domain-row';
      const tagLabel = document.createElement('label');
      tagLabel.className = 'dice-domain-label';
      tagLabel.textContent = 'Equipment tag modifier:';
      equipmentTagSelect = document.createElement('select');
      [
        { value: '0', label: 'None (0)' },
        { value: '1', label: 'Advantage (+1 die)' },
        { value: '-1', label: 'Drawback (-1 die)' }
      ].forEach((optData) => {
        const opt = document.createElement('option');
        opt.value = optData.value;
        opt.textContent = optData.label;
        equipmentTagSelect.appendChild(opt);
      });
      tagRow.appendChild(tagLabel);
      tagRow.appendChild(equipmentTagSelect);
      wrap.appendChild(tagRow);
    }

    // Result display
    const resultEl = document.createElement('div');
    resultEl.className = 'dice-result';
    resultEl.innerHTML = '<span class="dice-result-prompt">Roll a die above</span>';
    wrap.appendChild(resultEl);

    // History
    const histTitle = document.createElement('div');
    histTitle.className = 'dice-hist-title';
    histTitle.textContent = 'Roll history';
    wrap.appendChild(histTitle);
    const histList = document.createElement('ul');
    histList.className = 'dice-hist-list';
    histList.id = 'dice-hist-list';
    wrap.appendChild(histList);

    content.appendChild(wrap);

    function rollDie(sides, label, el) {
      const rules = getRulesConfig();
      const mod = parseInt(document.getElementById('dice-difficulty').value, 10);
      const base = composerMasteredChk ? (composerMasteredChk.checked ? 2 : 1) : (poolBonus || 1);
      const domainBonus = composerDomainChk
        ? (composerDomainChk.checked ? 1 : 0)
        : (domainBonusChk && domainBonusChk.checked ? 1 : 0);
      const equipmentBonus = composerTagImpact
        ? (parseInt(composerTagImpact.value, 10) || 0)
        : (equipmentTagSelect ? (parseInt(equipmentTagSelect.value, 10) || 0) : 0);
      const pool = base + mod + domainBonus + equipmentBonus;
      const rollCount = Math.max(1, pool);
      const downgradeSteps = rules.difficultyDowngrades ? Math.max(0, -pool) : 0;
      let rolls = [];
      for (let i = 0; i < rollCount; i++) rolls.push(Math.ceil(Math.random() * sides));
      const kept = Math.max(...rolls);

      el.innerHTML = '';
      const big = document.createElement('div');
      big.className = 'dice-big-result';
      big.textContent = kept;
      const sub = document.createElement('div');
      sub.className = 'dice-sub';

      let note = rollCount > 1 ? 'kept highest' : '';
      if (downgradeSteps > 0) note = (note ? note + ', ' : '') + `outcome downgraded ${downgradeSteps} step${downgradeSteps > 1 ? 's' : ''}`;
      if (equipmentBonus > 0) note = (note ? note + ', ' : '') + 'equipment advantage';
      if (equipmentBonus < 0) note = (note ? note + ', ' : '') + 'equipment drawback';

      if (sides === 10) {
        const rawBand = d10OutcomeBand(kept);
        const finalBand = downgradeOutcomeBand(rawBand.index, downgradeSteps);
        sub.textContent = `${label} — ${rolls.join(', ')} | ${rawBand.label} → ${finalBand.label}`;
      } else {
        sub.textContent = rollCount > 1 ? `${label} — rolled ${rolls.join(', ')}` : label;
      }
      if (note) sub.textContent += ` (${note})`;

      el.appendChild(big);
      el.appendChild(sub);
      addToHistory(histList, label, kept, rolls, note, sides, downgradeSteps);
      if (options.logTargetId) appendLog(`Rolled ${label}: ${kept}`, options.logTargetId);
    }

    function d10OutcomeBand(value) {
      if (value <= 1) return { index: 0, label: 'Critical Failure (1)' };
      if (value <= 5) return { index: 1, label: 'Failure (2-5)' };
      if (value <= 7) return { index: 2, label: 'Success with Cost (6-7)' };
      if (value <= 9) return { index: 3, label: 'Success (8-9)' };
      return { index: 4, label: 'Critical Success (10)' };
    }

    function downgradeOutcomeBand(index, steps) {
      const labels = ['Critical Failure (1)', 'Failure (2-5)', 'Success with Cost (6-7)', 'Success (8-9)', 'Critical Success (10)'];
      const clamped = Math.max(0, Math.min(labels.length - 1, index - steps));
      return { index: clamped, label: labels[clamped] };
    }

    function addToHistory(list, label, result, rolls, note, sides, downgradeSteps) {
      const li = document.createElement('li');
      if (sides === 10) {
        const rawBand = d10OutcomeBand(result);
        const finalBand = downgradeOutcomeBand(rawBand.index, downgradeSteps || 0);
        li.textContent = `${label}: ${result} (${rolls.join(', ')} — ${rawBand.label} -> ${finalBand.label}${note ? ', ' + note : ''})`;
      } else {
        li.textContent = label + ': ' + result + (note ? ' (' + rolls.join(', ') + ' — ' + note + ')' : '');
      }
      list.insertBefore(li, list.firstChild);
      if (list.children.length > 8) list.removeChild(list.lastChild);
    }
  }

  // -----------------------------------------------------------------------
  //  SETTINGS PANEL
  // -----------------------------------------------------------------------
  function openSettingsModal() {
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    const content = document.getElementById('modal-content');
    const titleEl = document.getElementById('modal-title');
    if (titleEl) titleEl.textContent = 'Settings';
    content.innerHTML = '';
    overlay.classList.remove('hidden');
    modal.classList.remove('hidden');

    const camp = currentCampaign();

    // Campaign name
    const nameField = document.createElement('div');
    nameField.className = 'modal-field';
    nameField.innerHTML = '<label>Campaign Name</label>';
    const nameInput = document.createElement('input');
    nameInput.type = 'text';
    nameInput.value = camp.name;
    nameInput.addEventListener('change', e => {
      camp.name = e.target.value.trim() || camp.name;
      renderCampaignName();
      saveCampaigns();
    });
    nameField.appendChild(nameInput);
    content.appendChild(nameField);

    // GM PIN
    const pinField = document.createElement('div');
    pinField.className = 'modal-field';
    pinField.innerHTML = '<label>GM PIN (leave blank for no protection)</label>';
    const pinInput = document.createElement('input');
    pinInput.type = 'password';
    pinInput.value = camp.gmPin || '';
    pinInput.placeholder = 'e.g. 1234';
    pinInput.maxLength = 8;
    pinInput.addEventListener('change', e => {
      camp.gmPin = e.target.value.trim();
      saveCampaigns();
      showToast('GM PIN updated', 'info');
    });
    pinField.appendChild(pinInput);
    content.appendChild(pinField);

    // Player editing toggle
    const editField = document.createElement('div');
    editField.className = 'modal-field modal-field-inline';
    const editLabel = document.createElement('label');
    editLabel.textContent = 'Allow players to edit their own PC';
    const editChk = document.createElement('input');
    editChk.type = 'checkbox';
    editChk.checked = camp.allowPlayerEditing !== false;
    editChk.addEventListener('change', e => {
      camp.allowPlayerEditing = e.target.checked;
      saveCampaigns();
    });
    editField.appendChild(editChk);
    editField.appendChild(editLabel);
    content.appendChild(editField);

    // Collaboration (invite + permissions)
    if (!Array.isArray(camp.memberUsers)) camp.memberUsers = [];
    if (camp.owner && !camp.memberUsers.includes(camp.owner)) camp.memberUsers.unshift(camp.owner);
    const collabField = document.createElement('div');
    collabField.className = 'modal-field';
    collabField.style.border = '1px solid var(--spire-border)';
    collabField.style.borderRadius = 'var(--radius)';
    collabField.style.padding = '8px';
    collabField.style.background = 'var(--spire-mid)';
    const collabLabel = document.createElement('label');
    collabLabel.textContent = 'Collaboration';
    collabField.appendChild(collabLabel);
    const campGms = Array.isArray(camp.gmUsers) ? camp.gmUsers : [];
    const canManageCollab = !!(state.currentUser && (camp.owner === state.currentUser || campGms.includes(state.currentUser)));

    // Pull latest members from shared invite snapshot to avoid overwriting joins.
    if (camp.inviteCode) {
      const sharedRec = loadSharedInvites()[String(camp.inviteCode).trim().toUpperCase()];
      const sharedMembers = Array.isArray(sharedRec?.data?.memberUsers) ? sharedRec.data.memberUsers : [];
      if (sharedMembers.length) {
        camp.memberUsers = Array.from(new Set([...(camp.memberUsers || []), ...sharedMembers]));
      }
    }

    const inviteRow = document.createElement('div');
    inviteRow.style.display = 'flex';
    inviteRow.style.gap = '6px';
    inviteRow.style.alignItems = 'center';
    inviteRow.style.marginTop = '6px';
    const inviteInput = document.createElement('input');
    inviteInput.type = 'text';
    inviteInput.readOnly = true;
    inviteInput.value = camp.inviteCode || '';
    inviteInput.placeholder = 'No active invite code';
    inviteInput.style.flex = '1';
    inviteRow.appendChild(inviteInput);

    const copyInviteBtn = document.createElement('button');
    copyInviteBtn.textContent = 'Copy';
    copyInviteBtn.disabled = !camp.inviteCode;
    copyInviteBtn.addEventListener('click', async () => {
      if (!camp.inviteCode) return;
      try {
        if (navigator.clipboard?.writeText) await navigator.clipboard.writeText(camp.inviteCode);
        else {
          inviteInput.select();
          document.execCommand('copy');
        }
        showToast('Invite code copied.', 'info');
      } catch (_) {
        showToast('Could not copy invite code.', 'warn');
      }
    });
    inviteRow.appendChild(copyInviteBtn);

    const genInviteBtn = document.createElement('button');
    genInviteBtn.textContent = camp.inviteCode ? 'Rotate' : 'Generate';
    genInviteBtn.disabled = !canManageCollab;
    genInviteBtn.addEventListener('click', async () => {
      if (canUseOnlineApi() && initOnlineClient()) {
        try {
          const cloudId = camp.cloudCampaignId || camp.id;
          const code = await window.SpireOnlineClient.generateInviteCode(cloudId, 'player', 1, 1440);
          const normalizedCode = String(code || '').trim().toUpperCase();
          if (!normalizedCode) {
            showToast('Invite generated but no code was returned.', 'warn');
            return;
          }
          camp.cloudCampaignId = cloudId;
          camp.inviteCode = normalizedCode;
          saveCampaigns();
          openSettingsModal();
          showToast('Cloud invite code generated.', 'info');
          return;
        } catch (err) {
          showToast(err?.message || 'Could not generate cloud invite.', 'warn');
          return;
        }
      }
      camp.inviteCode = generateInviteCode();
      publishCampaignInvite(camp);
      saveCampaigns();
      openSettingsModal();
      showToast('Invite code updated.', 'info');
    });
    inviteRow.appendChild(genInviteBtn);

    const revokeInviteBtn = document.createElement('button');
    revokeInviteBtn.textContent = 'Revoke';
    revokeInviteBtn.disabled = !canManageCollab || !camp.inviteCode;
    revokeInviteBtn.addEventListener('click', async () => {
      if (!camp.inviteCode) return;
      const ok = await askConfirm('Revoke this invite code?', 'Revoke Invite');
      if (!ok) return;
      if (canUseOnlineApi() && initOnlineClient()) {
        try {
          await window.SpireOnlineClient.revokeInviteCode(camp.cloudCampaignId || camp.id, camp.inviteCode);
        } catch (err) {
          showToast(err?.message || 'Could not revoke cloud invite.', 'warn');
          return;
        }
      }
      revokeCampaignInvite(camp);
      camp.inviteCode = '';
      saveCampaigns();
      openSettingsModal();
      showToast('Invite code revoked.', 'info');
    });
    inviteRow.appendChild(revokeInviteBtn);
    collabField.appendChild(inviteRow);

    const collabMeta = document.createElement('div');
    collabMeta.className = 'text-muted';
    collabMeta.style.fontSize = '0.8rem';
    collabMeta.style.marginTop = '4px';
    collabMeta.textContent = canManageCollab
      ? 'Share invite code with players. GMs can manage invite and membership.'
      : `Campaign owner: ${camp.owner || 'Unknown'}. Ask a GM to generate or revoke invite codes.`;
    collabField.appendChild(collabMeta);

    const membersWrap = document.createElement('div');
    membersWrap.style.display = 'flex';
    membersWrap.style.flexDirection = 'column';
    membersWrap.style.gap = '6px';
    membersWrap.style.marginTop = '8px';
    const membersHeader = document.createElement('strong');
    membersHeader.textContent = 'Members';
    membersWrap.appendChild(membersHeader);

    const users = camp.memberUsers.length ? camp.memberUsers.slice() : (camp.owner ? [camp.owner] : []);
    users.forEach((username) => {
      const row = document.createElement('div');
      row.className = 'settings-rel-row';
      const role = username === camp.owner ? 'Owner' : (camp.gmUsers || []).includes(username) ? 'Co-GM' : 'Player';
      const label = document.createElement('span');
      label.textContent = `${username} (${role})`;
      row.appendChild(label);

      if (canManageCollab && username !== camp.owner) {
        const toggleRoleBtn = document.createElement('button');
        toggleRoleBtn.textContent = (camp.gmUsers || []).includes(username) ? 'Demote' : 'Promote GM';
        toggleRoleBtn.style.marginRight = '4px';
        toggleRoleBtn.addEventListener('click', () => {
          if (!Array.isArray(camp.gmUsers)) camp.gmUsers = [];
          if (camp.gmUsers.includes(username)) camp.gmUsers = camp.gmUsers.filter((u) => u !== username);
          else camp.gmUsers.push(username);
          saveCampaigns();
          openSettingsModal();
        });
        row.appendChild(toggleRoleBtn);

        const removeBtn = document.createElement('button');
        removeBtn.className = 'row-remove-btn';
        removeBtn.textContent = '×';
        removeBtn.title = 'Remove member';
        removeBtn.addEventListener('click', () => {
          camp.memberUsers = (camp.memberUsers || []).filter((u) => u !== username);
          camp.gmUsers = (camp.gmUsers || []).filter((u) => u !== username);
          saveCampaigns();
          openSettingsModal();
        });
        row.appendChild(removeBtn);
      }
      membersWrap.appendChild(row);
    });

    if (canManageCollab) {
      const addRow = document.createElement('div');
      addRow.style.display = 'flex';
      addRow.style.gap = '6px';
      const addSel = document.createElement('select');
      const knownUsers = Object.keys(state.users || {})
        .filter((u) => !(camp.memberUsers || []).includes(u))
        .sort((a, b) => a.localeCompare(b));
      const emptyOpt = document.createElement('option');
      emptyOpt.value = '';
      emptyOpt.textContent = knownUsers.length ? 'Add user…' : 'No available users';
      addSel.appendChild(emptyOpt);
      knownUsers.forEach((u) => {
        const opt = document.createElement('option');
        opt.value = u;
        opt.textContent = u;
        addSel.appendChild(opt);
      });
      addRow.appendChild(addSel);
      const addBtn = document.createElement('button');
      addBtn.textContent = 'Add';
      addBtn.disabled = !knownUsers.length;
      addBtn.addEventListener('click', () => {
        if (!addSel.value) return;
        if (!Array.isArray(camp.memberUsers)) camp.memberUsers = [];
        if (!camp.memberUsers.includes(addSel.value)) camp.memberUsers.push(addSel.value);
        saveCampaigns();
        openSettingsModal();
      });
      addRow.appendChild(addBtn);
      membersWrap.appendChild(addRow);
    }

    collabField.appendChild(membersWrap);
    content.appendChild(collabField);

    // Sync transport (local channel vs Supabase realtime)
    const syncCfg = state.syncConfig || loadSyncConfig();
    const syncField = document.createElement('div');
    syncField.className = 'modal-field';
    syncField.style.border = '1px solid var(--spire-border)';
    syncField.style.borderRadius = 'var(--radius)';
    syncField.style.padding = '8px';
    syncField.style.background = 'var(--spire-mid)';
    const syncLabel = document.createElement('label');
    syncLabel.textContent = 'Realtime Transport';
    syncField.appendChild(syncLabel);

    const transportSel = document.createElement('select');
    [
      { value: 'local', label: 'Local (multi-tab on same browser)' },
      { value: 'supabase', label: 'Supabase Realtime (cross-device)' }
    ].forEach((optData) => {
      const opt = document.createElement('option');
      opt.value = optData.value;
      opt.textContent = optData.label;
      if ((syncCfg.transport || 'local') === optData.value) opt.selected = true;
      transportSel.appendChild(opt);
    });
    transportSel.style.marginTop = '6px';
    syncField.appendChild(transportSel);

    const syncMeta = document.createElement('div');
    syncMeta.className = 'text-muted';
    syncMeta.style.fontSize = '0.78rem';
    syncMeta.style.marginTop = '4px';
    syncField.appendChild(syncMeta);

    const supaWrap = document.createElement('div');
    supaWrap.style.display = 'grid';
    supaWrap.style.gridTemplateColumns = '1fr';
    supaWrap.style.gap = '6px';
    supaWrap.style.marginTop = '8px';
    const supaUrl = document.createElement('input');
    supaUrl.type = 'text';
    supaUrl.placeholder = 'Supabase URL';
    supaUrl.value = syncCfg.supabaseUrl || '';
    const supaKey = document.createElement('input');
    supaKey.type = 'password';
    supaKey.placeholder = 'Supabase anon key';
    supaKey.value = syncCfg.supabaseAnonKey || '';
    supaWrap.appendChild(supaUrl);
    supaWrap.appendChild(supaKey);
    syncField.appendChild(supaWrap);

    const reconnectBtn = document.createElement('button');
    reconnectBtn.textContent = 'Reconnect Realtime';
    reconnectBtn.style.marginTop = '8px';
    reconnectBtn.addEventListener('click', () => {
      const nextCfg = saveSyncConfig({
        transport: transportSel.value,
        supabaseUrl: supaUrl.value,
        supabaseAnonKey: supaKey.value
      });
      initRealtimeChannel();
      const active = state.realtimeTransport === 'supabase' && nextCfg.transport === 'supabase'
        ? 'Supabase realtime connected.'
        : `Realtime transport: ${state.realtimeTransport}.`;
      showToast(active, 'info');
      openSettingsModal();
    });
    syncField.appendChild(reconnectBtn);

    function syncTransportVisibility() {
      const wantsSupa = transportSel.value === 'supabase';
      supaWrap.style.display = wantsSupa ? 'grid' : 'none';
      syncMeta.textContent = wantsSupa
        ? 'Requires Supabase URL + anon key and online-client setup.'
        : 'Uses BroadcastChannel for same-browser tab sync.';
    }
    syncTransportVisibility();
    transportSel.addEventListener('change', () => {
      syncTransportVisibility();
      saveSyncConfig({
        transport: transportSel.value,
        supabaseUrl: supaUrl.value,
        supabaseAnonKey: supaKey.value
      });
      initRealtimeChannel();
      updateSyncQueueButton();
    });
    supaUrl.addEventListener('change', () => {
      saveSyncConfig({ transport: transportSel.value, supabaseUrl: supaUrl.value, supabaseAnonKey: supaKey.value });
      initRealtimeChannel();
    });
    supaKey.addEventListener('change', () => {
      saveSyncConfig({ transport: transportSel.value, supabaseUrl: supaUrl.value, supabaseAnonKey: supaKey.value });
      initRealtimeChannel();
    });
    content.appendChild(syncField);

    // Scenario packs / plugin scaffold
    if (!Array.isArray(camp.scenarioPacks)) camp.scenarioPacks = [];
    const packField = document.createElement('div');
    packField.className = 'modal-field';
    packField.style.border = '1px solid var(--spire-border)';
    packField.style.borderRadius = 'var(--radius)';
    packField.style.padding = '8px';
    packField.style.background = 'var(--spire-mid)';
    const packLabel = document.createElement('label');
    packLabel.textContent = 'Scenario Packs (Plugin Scaffold)';
    packField.appendChild(packLabel);
    const packMeta = document.createElement('div');
    packMeta.className = 'text-muted';
    packMeta.style.fontSize = '0.78rem';
    packMeta.style.marginBottom = '6px';
    packMeta.textContent = 'Import JSON packs to extend NPC templates and scene prompt tables without editing core code.';
    packField.appendChild(packMeta);
    const packList = document.createElement('div');
    packList.className = 'sync-queue-list';
    packField.appendChild(packList);
    const packInput = document.createElement('input');
    packInput.type = 'file';
    packInput.accept = 'application/json';
    packInput.style.display = 'none';
    packInput.addEventListener('change', (e) => {
      const file = e.target.files && e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const parsed = JSON.parse(String(reader.result || '{}'));
          const normalized = normalizeScenarioPack(parsed);
          if (!normalized) {
            showToast('Invalid pack format.', 'warn');
            return;
          }
          const existing = (camp.scenarioPacks || []).find((p) => p.id === normalized.id || p.name.toLowerCase() === normalized.name.toLowerCase());
          if (existing) {
            Object.assign(existing, normalized);
          } else {
            camp.scenarioPacks.push(normalized);
          }
          appendLog('Imported scenario pack', '');
          saveCampaigns();
          openSettingsModal();
          showToast(`Pack "${normalized.name}" loaded.`, 'info');
        } catch (_) {
          showToast('Pack import failed.', 'warn');
        } finally {
          e.target.value = '';
        }
      };
      reader.readAsText(file);
    });
    packField.appendChild(packInput);
    const packActions = document.createElement('div');
    packActions.style.display = 'flex';
    packActions.style.gap = '6px';
    packActions.style.marginTop = '6px';
    const importPackBtn = document.createElement('button');
    importPackBtn.textContent = 'Import Pack';
    importPackBtn.addEventListener('click', () => packInput.click());
    packActions.appendChild(importPackBtn);
    const templateBtn = document.createElement('button');
    templateBtn.textContent = 'Download Template';
    templateBtn.addEventListener('click', () => {
      const sample = {
        id: 'pack-sample',
        name: 'Sample Pack',
        source: 'local',
        enabled: true,
        scenePrompts: {
          complications: ['A shrine has been sealed by unknown agents.'],
          factionReactions: ['A hidden ministry proxy starts asking questions.'],
          twists: ['A trusted intermediary is an impostor.']
        },
        npcTemplates: {
          ministry_agent: {
            label: 'Ministry Agent',
            role: 'Ministry Agent',
            threatLevel: 'Significant',
            disposition: 'Wary',
            wants: 'Contain dissent quickly',
            fears: 'Political embarrassment',
            leverage: 'Official authority and informant network',
            inventory: [
              { item: 'Service blade', quantity: 1, type: 'weapon', stress: 'D6 stress', tags: ['Parrying'] }
            ]
          }
        }
      };
      const blob = new Blob([JSON.stringify(sample, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'scenario-pack-template.json';
      link.click();
      URL.revokeObjectURL(url);
    });
    packActions.appendChild(templateBtn);
    packField.appendChild(packActions);
    if (!camp.scenarioPacks.length) {
      const empty = document.createElement('div');
      empty.className = 'text-muted';
      empty.style.fontSize = '0.8rem';
      empty.style.padding = '6px';
      empty.textContent = 'No packs installed.';
      packList.appendChild(empty);
    } else {
      camp.scenarioPacks.forEach((pack) => {
        const row = document.createElement('div');
        row.className = 'sync-queue-row';
        const chk = document.createElement('input');
        chk.type = 'checkbox';
        chk.checked = pack.enabled !== false;
        chk.addEventListener('change', () => {
          pack.enabled = !!chk.checked;
          saveCampaigns();
        });
        const body = document.createElement('div');
        body.className = 'sync-queue-row-body';
        const nameEl = document.createElement('div');
        nameEl.className = 'sync-queue-when';
        nameEl.textContent = pack.name || 'Scenario Pack';
        const metaEl = document.createElement('div');
        metaEl.className = 'sync-queue-meta';
        const promptCount = ((pack.scenePrompts?.complications || []).length + (pack.scenePrompts?.factionReactions || []).length + (pack.scenePrompts?.twists || []).length);
        const templateCount = Object.keys(pack.npcTemplates || {}).length;
        metaEl.textContent = `${templateCount} template(s), ${promptCount} prompts`;
        body.appendChild(nameEl);
        body.appendChild(metaEl);
        const removeBtn = document.createElement('button');
        removeBtn.className = 'row-remove-btn';
        removeBtn.textContent = '×';
        removeBtn.title = 'Remove pack';
        removeBtn.addEventListener('click', () => {
          camp.scenarioPacks = camp.scenarioPacks.filter((p) => p.id !== pack.id);
          appendLog('Removed scenario pack', '');
          saveCampaigns();
          openSettingsModal();
        });
        row.appendChild(chk);
        row.appendChild(body);
        row.appendChild(removeBtn);
        packList.appendChild(row);
      });
    }
    content.appendChild(packField);

    // Crash reporting / telemetry
    const teleCfg = state.telemetryConfig || loadTelemetryConfig();
    const teleField = document.createElement('div');
    teleField.className = 'modal-field';
    teleField.style.border = '1px solid var(--spire-border)';
    teleField.style.borderRadius = 'var(--radius)';
    teleField.style.padding = '8px';
    teleField.style.background = 'var(--spire-mid)';
    const teleLabel = document.createElement('label');
    teleLabel.textContent = 'Crash Reporting';
    teleField.appendChild(teleLabel);
    const teleToggleRow = document.createElement('div');
    teleToggleRow.className = 'modal-field modal-field-inline';
    const teleToggle = document.createElement('input');
    teleToggle.type = 'checkbox';
    teleToggle.checked = !!teleCfg.enabled;
    const teleToggleLabel = document.createElement('label');
    teleToggleLabel.textContent = 'Enable crash reporting';
    teleToggleRow.appendChild(teleToggle);
    teleToggleRow.appendChild(teleToggleLabel);
    teleField.appendChild(teleToggleRow);
    const teleEndpoint = document.createElement('input');
    teleEndpoint.type = 'text';
    teleEndpoint.placeholder = 'Webhook endpoint (optional)';
    teleEndpoint.value = teleCfg.endpoint || '';
    teleField.appendChild(teleEndpoint);
    const teleMeta = document.createElement('div');
    teleMeta.className = 'text-muted';
    teleMeta.style.fontSize = '0.78rem';
    const localCrashCount = loadCrashLog().length;
    teleMeta.textContent = `Local crash log entries: ${localCrashCount}`;
    teleField.appendChild(teleMeta);
    const teleActions = document.createElement('div');
    teleActions.style.display = 'flex';
    teleActions.style.gap = '6px';
    const saveTeleBtn = document.createElement('button');
    saveTeleBtn.textContent = 'Save Telemetry';
    saveTeleBtn.addEventListener('click', () => {
      saveTelemetryConfig({
        enabled: teleToggle.checked,
        endpoint: teleEndpoint.value,
        maxLocal: teleCfg.maxLocal || 50
      });
      showToast('Crash reporting settings saved.', 'info');
      openSettingsModal();
    });
    teleActions.appendChild(saveTeleBtn);
    const exportTeleBtn = document.createElement('button');
    exportTeleBtn.textContent = 'Export Crash Log';
    exportTeleBtn.addEventListener('click', () => {
      const data = loadCrashLog();
      const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `spire-crash-log-${Date.now()}.json`;
      link.click();
      URL.revokeObjectURL(url);
    });
    teleActions.appendChild(exportTeleBtn);
    const clearTeleBtn = document.createElement('button');
    clearTeleBtn.textContent = 'Clear Crash Log';
    clearTeleBtn.addEventListener('click', async () => {
      const ok = await askConfirm('Clear local crash log entries?', 'Clear Crash Log');
      if (!ok) return;
      saveCrashLog([]);
      showToast('Crash log cleared.', 'info');
      openSettingsModal();
    });
    teleActions.appendChild(clearTeleBtn);
    teleField.appendChild(teleActions);
    content.appendChild(teleField);

    // Rules profile
    const rulesField = document.createElement('div');
    rulesField.className = 'modal-field';
    const rulesLabel = document.createElement('label');
    rulesLabel.textContent = 'Rules Profile';
    const rulesSel = document.createElement('select');
    [
      { value: 'Core', label: 'Core (recommended)' },
      { value: 'Quickstart', label: 'Quickstart (lighter bookkeeping)' },
      { value: 'Custom', label: 'Custom' }
    ].forEach(optData => {
      const opt = document.createElement('option');
      opt.value = optData.value;
      opt.textContent = optData.label;
      if ((camp.rulesProfile || 'Core') === optData.value) opt.selected = true;
      rulesSel.appendChild(opt);
    });
    rulesField.appendChild(rulesLabel);
    rulesField.appendChild(rulesSel);
    content.appendChild(rulesField);

    // Rules preset diff panel: explicit view of what each profile changes.
    const rulesDiffField = document.createElement('div');
    rulesDiffField.className = 'modal-field';
    rulesDiffField.style.border = '1px solid var(--spire-border)';
    rulesDiffField.style.borderRadius = 'var(--radius)';
    rulesDiffField.style.padding = '8px';
    rulesDiffField.style.background = 'var(--spire-mid)';
    const diffTitle = document.createElement('label');
    diffTitle.textContent = 'Preset Differences';
    rulesDiffField.appendChild(diffTitle);
    const diffGrid = document.createElement('div');
    diffGrid.style.display = 'grid';
    diffGrid.style.gridTemplateColumns = '1.4fr 1fr 1fr 1fr';
    diffGrid.style.gap = '6px';
    diffGrid.style.fontSize = '0.78rem';
    rulesDiffField.appendChild(diffGrid);
    content.appendChild(rulesDiffField);

    // Custom rules toggles (visible only when profile = Custom)
    if (!camp.customRules) {
      camp.customRules = {
        difficultyDowngrades: true,
        falloutCheckOnStress: true,
        clearStressOnFallout: true
      };
    }
    const customRulesField = document.createElement('div');
    customRulesField.className = 'modal-field';
    customRulesField.style.border = '1px solid var(--spire-border)';
    customRulesField.style.borderRadius = 'var(--radius)';
    customRulesField.style.padding = '8px';
    customRulesField.style.background = 'var(--spire-mid)';

    function makeRuleToggle(labelText, key) {
      const row = document.createElement('div');
      row.className = 'modal-field modal-field-inline';
      row.style.marginBottom = '6px';
      const chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.checked = camp.customRules[key] !== false;
      chk.addEventListener('change', e => {
        camp.customRules[key] = e.target.checked;
        saveCampaigns();
        renderRulesDiff();
      });
      const lbl = document.createElement('label');
      lbl.textContent = labelText;
      row.appendChild(chk);
      row.appendChild(lbl);
      return row;
    }

    customRulesField.appendChild(makeRuleToggle('Use difficulty outcome downgrades at negative pool', 'difficultyDowngrades'));
    customRulesField.appendChild(makeRuleToggle('Check Fallout when stress is taken', 'falloutCheckOnStress'));
    customRulesField.appendChild(makeRuleToggle('Clear stress automatically when Fallout triggers', 'clearStressOnFallout'));
    content.appendChild(customRulesField);

    function renderRulesDiff() {
      diffGrid.innerHTML = '';
      const headers = ['Rule', 'Core', 'Quickstart', 'Custom'];
      headers.forEach((h, i) => {
        const cell = document.createElement('div');
        cell.textContent = h;
        cell.style.fontWeight = '700';
        cell.style.color = 'var(--accent-hi)';
        if (i > 0 && rulesSel.value.toLowerCase() === h.toLowerCase()) {
          cell.style.textDecoration = 'underline';
        }
        diffGrid.appendChild(cell);
      });
      const rows = [
        {
          label: 'Difficulty downgrades',
          core: true,
          quick: true,
          custom: camp.customRules.difficultyDowngrades !== false
        },
        {
          label: 'Fallout check on stress',
          core: true,
          quick: false,
          custom: camp.customRules.falloutCheckOnStress !== false
        },
        {
          label: 'Auto-clear stress on fallout',
          core: true,
          quick: false,
          custom: camp.customRules.clearStressOnFallout !== false
        }
      ];
      rows.forEach((row) => {
        const labelCell = document.createElement('div');
        labelCell.textContent = row.label;
        diffGrid.appendChild(labelCell);
        [row.core, row.quick, row.custom].forEach((v) => {
          const valCell = document.createElement('div');
          valCell.textContent = v ? 'Enabled' : 'Disabled';
          valCell.style.color = v ? 'var(--ok)' : 'var(--spire-dim)';
          diffGrid.appendChild(valCell);
        });
      });
    }

    function syncRulesVisibility() {
      customRulesField.style.display = (rulesSel.value === 'Custom') ? 'block' : 'none';
      renderRulesDiff();
    }
    syncRulesVisibility();
    rulesSel.addEventListener('change', e => {
      camp.rulesProfile = e.target.value;
      saveCampaigns();
      syncRulesVisibility();
    });

    // Fallout lookup prompt editor (campaign-specific)
    if (!camp.falloutGuidance) camp.falloutGuidance = defaultFalloutGuidance();
    const falloutSettings = document.createElement('div');
    falloutSettings.className = 'modal-field';
    const falloutLabel = document.createElement('label');
    falloutLabel.textContent = 'Fallout Lookup Prompts';
    const falloutMeta = document.createElement('div');
    falloutMeta.className = 'text-muted';
    falloutMeta.style.fontSize = '0.8rem';
    falloutMeta.style.marginBottom = '6px';
    falloutMeta.textContent = 'Edit prompt lists used by the Fallout Lookup button (one prompt per line).';

    const falloutTrackSel = document.createElement('select');
    ['Blood','Mind','Silver','Shadow','Reputation'].forEach((track) => {
      const opt = document.createElement('option');
      opt.value = track;
      opt.textContent = track;
      falloutTrackSel.appendChild(opt);
    });
    falloutTrackSel.style.marginBottom = '8px';

    const falloutEditorWrap = document.createElement('div');
    falloutEditorWrap.style.display = 'flex';
    falloutEditorWrap.style.flexDirection = 'column';
    falloutEditorWrap.style.gap = '8px';

    function renderFalloutEditor(track) {
      falloutEditorWrap.innerHTML = '';
      if (!camp.falloutGuidance[track]) {
        camp.falloutGuidance[track] = { Minor: [], Moderate: [], Severe: [] };
      }
      ['Minor', 'Moderate', 'Severe'].forEach((severity) => {
        const row = document.createElement('div');
        row.className = 'modal-field';
        row.style.marginBottom = '0';
        const lbl = document.createElement('label');
        lbl.textContent = severity;
        const ta = document.createElement('textarea');
        ta.rows = 3;
        ta.value = (camp.falloutGuidance[track][severity] || []).join('\n');
        ta.addEventListener('change', (e) => {
          camp.falloutGuidance[track][severity] = e.target.value
            .split('\n')
            .map(v => v.trim())
            .filter(Boolean);
          saveCampaigns();
        });
        row.appendChild(lbl);
        row.appendChild(ta);
        falloutEditorWrap.appendChild(row);
      });
    }

    falloutTrackSel.addEventListener('change', () => {
      renderFalloutEditor(falloutTrackSel.value);
    });
    renderFalloutEditor('Blood');

    const resetFalloutBtn = document.createElement('button');
    resetFalloutBtn.textContent = 'Reset Fallout Prompts to Defaults';
    resetFalloutBtn.addEventListener('click', async () => {
      const ok = await askConfirm('Reset all fallout lookup prompts in this campaign to defaults?', 'Reset Fallout Prompts');
      if (!ok) return;
      camp.falloutGuidance = defaultFalloutGuidance();
      renderFalloutEditor(falloutTrackSel.value || 'Blood');
      saveCampaigns();
      showToast('Fallout prompts reset to defaults.', 'info');
    });

    falloutSettings.appendChild(falloutLabel);
    falloutSettings.appendChild(falloutMeta);
    falloutSettings.appendChild(falloutTrackSel);
    falloutSettings.appendChild(falloutEditorWrap);
    falloutSettings.appendChild(resetFalloutBtn);
    content.appendChild(falloutSettings);

    // Relationship types
    const relSection = document.createElement('div');
    relSection.className = 'modal-field';
    relSection.innerHTML = '<label>Relationship Types</label>';
    const relList = document.createElement('div');
    relList.className = 'settings-rel-list';
    function renderRelList() {
      relList.innerHTML = '';
      camp.relTypes.forEach((rt, i) => {
        const row = document.createElement('div');
        row.className = 'settings-rel-row';
        const span = document.createElement('span');
        span.textContent = rt;
        const del = document.createElement('button');
        del.textContent = '×';
        del.className = 'row-remove-btn';
        del.title = 'Remove this type';
        del.addEventListener('click', () => {
          camp.relTypes.splice(i, 1);
          state.relTypes = camp.relTypes;
          saveCampaigns();
          renderRelList();
        });
        row.appendChild(span);
        row.appendChild(del);
        relList.appendChild(row);
      });
    }
    renderRelList();
    relSection.appendChild(relList);
    const addRelRow = document.createElement('div');
    addRelRow.style.display = 'flex';
    addRelRow.style.gap = '6px';
    addRelRow.style.marginTop = '6px';
    const newRelInput = document.createElement('input');
    newRelInput.type = 'text';
    newRelInput.placeholder = 'New type…';
    const addRelBtn = document.createElement('button');
    addRelBtn.textContent = 'Add';
    addRelBtn.addEventListener('click', () => {
      const val = newRelInput.value.trim();
      if (val && !camp.relTypes.includes(val)) {
        camp.relTypes.push(val);
        state.relTypes = camp.relTypes;
        saveCampaigns();
        renderRelList();
        newRelInput.value = '';
      }
    });
    addRelRow.appendChild(newRelInput);
    addRelRow.appendChild(addRelBtn);
    relSection.appendChild(addRelRow);
    content.appendChild(relSection);

    // Backup restore
    const backupRaw = localStorage.getItem(userScopedKey('spire-campaigns-backup'));
    const backupTs = localStorage.getItem(userScopedKey('spire-campaigns-backup-ts'));
    if (backupRaw) {
      const backupField = document.createElement('div');
      backupField.className = 'modal-field';
      const backupLabel = document.createElement('label');
      backupLabel.textContent = 'Recovery';
      const backupMeta = document.createElement('div');
      backupMeta.className = 'text-muted';
      backupMeta.style.fontSize = '0.8rem';
      backupMeta.style.marginBottom = '6px';
      backupMeta.textContent = 'Latest snapshot: ' + (backupTs ? new Date(backupTs).toLocaleString() : 'Unknown time');
      const restoreBtn = document.createElement('button');
      restoreBtn.textContent = 'Restore From Snapshot';
      restoreBtn.addEventListener('click', async () => {
        const ok = await askConfirm('Restore campaigns from the latest snapshot? This replaces current in-browser data.', 'Restore Snapshot');
        if (!ok) return;
        try {
          const parsed = JSON.parse(backupRaw);
          if (!parsed || !parsed.campaigns || !Object.keys(parsed.campaigns).length) {
            throw new Error('Backup snapshot is empty.');
          }
          state.campaigns = parsed.campaigns;
          state.currentCampaignId = parsed.currentCampaignId;
          if (!state.currentCampaignId || !state.campaigns[state.currentCampaignId]) {
            state.currentCampaignId = Object.keys(state.campaigns)[0];
          }
          closeModal();
          saveCampaigns();
          initAfterLoad();
          showToast('Restored data from snapshot.', 'warn');
        } catch (err) {
          console.error(err);
          showToast('Failed to restore snapshot.', 'warn');
        }
      });
      backupField.appendChild(backupLabel);
      backupField.appendChild(backupMeta);
      backupField.appendChild(restoreBtn);
      content.appendChild(backupField);
    }

    // Delete campaign
    const dangerZone = document.createElement('div');
    dangerZone.className = 'modal-field settings-danger';
    const deleteBtn = document.createElement('button');
    deleteBtn.className = 'pending-reject-btn';
    deleteBtn.style.width = '100%';
    deleteBtn.textContent = 'Delete Campaign';
    deleteBtn.addEventListener('click', async () => {
      const ok = await askConfirm('Permanently delete "' + camp.name + '"? This cannot be undone.', 'Delete Campaign');
      if (!ok) return;
      delete state.campaigns[camp.id];
      const remaining = Object.keys(state.campaigns);
      if (remaining.length === 0) {
        const fresh = createCampaign('New Campaign');
        state.campaigns[fresh.id] = fresh;
        state.currentCampaignId = fresh.id;
      } else {
        state.currentCampaignId = remaining[0];
      }
      saveCampaigns();
      closeModal();
      returnToModeScreen();
    });
    dangerZone.appendChild(deleteBtn);
    content.appendChild(dangerZone);
  }

  // -----------------------------------------------------------------------
  //  SESSION MANAGEMENT
  // -----------------------------------------------------------------------
  function startNewSession() {
    const camp = currentCampaign();
    camp.currentSession = (camp.currentSession || 1) + 1;
    // Reset core ability usage for all PCs
    Object.values(camp.entities).forEach(ent => {
      if (ent.type !== 'pc') return;
      if (ent.coreAbilitiesState) {
        Object.keys(ent.coreAbilitiesState).forEach(k => {
          ent.coreAbilitiesState[k] = false;
        });
      }
      ent.refreshed = false;
    });
    appendSessionLog('─── Session ' + camp.currentSession + ' begins ───');
    const badge = document.getElementById('log-session-badge');
    if (badge) badge.textContent = 'Session ' + camp.currentSession;
    showToast('Session ' + camp.currentSession + ' started. Core abilities and refresh flags reset.', 'info');
    saveAndRefresh();
    renderLog();
  }

  function openPlayerSubmitNPCModal() {
    const overlay = document.getElementById('modal-overlay');
    const modal = document.getElementById('modal');
    const content = document.getElementById('modal-content');
    const titleEl = document.getElementById('modal-title');
    if (titleEl) titleEl.textContent = 'Submit NPC Proposal';
    content.innerHTML = '';
    overlay.classList.remove('hidden');
    modal.classList.remove('hidden');

    const intro = document.createElement('p');
    intro.style.marginBottom = '12px';
    intro.style.color = 'var(--spire-muted)';
    intro.style.fontSize = '0.9rem';
    intro.textContent = 'Propose a new NPC. The GM will review and approve or reject your submission.';
    content.appendChild(intro);

    // Name field
    const nameField = document.createElement('div');
    nameField.className = 'modal-field';
    const nameLabel = document.createElement('label');
    nameLabel.textContent = 'Name';
    const nameInput = document.createElement('input');
    nameInput.type = 'text';
    nameInput.placeholder = 'NPC name…';
    nameField.appendChild(nameLabel);
    nameField.appendChild(nameInput);
    content.appendChild(nameField);

    // Role field
    const roleField = document.createElement('div');
    roleField.className = 'modal-field';
    const roleLabel = document.createElement('label');
    roleLabel.textContent = 'Role / Description';
    const roleInput = document.createElement('input');
    roleInput.type = 'text';
    roleInput.placeholder = 'e.g. Tavern keeper, Drow noble…';
    roleField.appendChild(roleLabel);
    roleField.appendChild(roleInput);
    content.appendChild(roleField);

    // Notes field
    const notesField = document.createElement('div');
    notesField.className = 'modal-field';
    const notesLabel = document.createElement('label');
    notesLabel.textContent = 'Notes (optional)';
    const notesInput = document.createElement('textarea');
    notesInput.placeholder = 'Any additional context for the GM…';
    notesInput.style.minHeight = '72px';
    notesField.appendChild(notesLabel);
    notesField.appendChild(notesInput);
    content.appendChild(notesField);

    // Buttons
    const btnBar = document.createElement('div');
    btnBar.style.display = 'flex';
    btnBar.style.gap = '8px';
    btnBar.style.marginTop = '12px';
    const submitBtn = document.createElement('button');
    submitBtn.className = 'modal-submit';
    submitBtn.textContent = 'Submit for Approval';
    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    cancelBtn.style.flex = '0 0 auto';
    btnBar.appendChild(submitBtn);
    btnBar.appendChild(cancelBtn);
    content.appendChild(btnBar);

    submitBtn.addEventListener('click', () => {
      const name = nameInput.value.trim();
      if (!name) { nameInput.focus(); return; }
      const camp = currentCampaign();
      const npc = newEntity(camp, 'npc');
      npc.name = name;
      npc.role = roleInput.value.trim() || 'NPC';
      npc.notes = notesInput.value.trim();
      npc.pendingApproval = true;
      appendLog('Submitted NPC for approval', npc.id);
      overlay.classList.add('hidden');
      modal.classList.add('hidden');
      saveAndRefresh();
      selectEntity(npc.id);
    });

    cancelBtn.addEventListener('click', () => {
      overlay.classList.add('hidden');
      modal.classList.add('hidden');
    });

    nameInput.focus();
  }

  /**
   * Add event listeners for UI elements and keyboard shortcuts.
   */
  function setupEventListeners() {
    if (eventListenersBound) return;
    eventListenersBound = true;
    window.addEventListener('storage', (e) => {
      if (!state.currentUser) return;
      const revKey = userScopedKey('spire-campaigns-rev');
      if (e.key !== revKey) return;
      if (!e.newValue) return;
      if (e.newValue === state.lastSeenRevision) return;
      const newlyConflicted = !state.syncConflictActive;
      state.lastSeenRevision = e.newValue;
      setSyncConflictWarning(true, conflictWarningText());
      setSaveState('error', conflictWarningText());
      if (newlyConflicted && (state.localEditsSinceConflict || 0) === 0) {
        setTimeout(async () => {
          if (!state.syncConflictActive || (state.localEditsSinceConflict || 0) > 0) return;
          const ok = await askConfirm(
            'Campaign updated in another tab. Reload latest now?',
            'Sync Conflict'
          );
          if (ok) reloadCampaignFromStorage();
        }, 60);
      }
    });
    window.addEventListener('online', () => {
      const flushed = flushPendingSaveOps();
      if (flushed) showToast('Back online. Queued changes synced.', 'info');
    });
    // Search box
    document.getElementById('search-input').addEventListener('input', () => {
      renderEntityLists();
    });

    /**
     * Open a name-picker modal for new entities. Provides buttons for
     * Drow-style names, modern names, or entering a custom name.
     * @param {string} type - 'pc', 'npc', or 'org'
     */
    function openNamePickerModal(type) {
      const overlay = document.getElementById('modal-overlay');
      const modal = document.getElementById('modal');
      const content = document.getElementById('modal-content');
      const titleEl = document.getElementById('modal-title');
      if (titleEl) titleEl.textContent = type === 'org' ? 'New Organisation' : `New ${type.toUpperCase()}`;
      content.innerHTML = '';
      overlay.classList.remove('hidden');
      modal.classList.remove('hidden');

      const intro = document.createElement('p');
      intro.style.marginBottom = '12px';
      intro.style.color = 'var(--spire-muted)';
      intro.textContent = type === 'org' ? 'Choose a name for the new organisation:' : 'Choose a name style:';
      content.appendChild(intro);
      let npcTemplateKey = 'none';
      if (type === 'npc') {
        const tplField = document.createElement('div');
        tplField.className = 'modal-field';
        const tplLabel = document.createElement('label');
        tplLabel.textContent = 'Quick Template';
        const tplSelect = document.createElement('select');
        const templates = getEffectiveNpcTemplates(currentCampaign());
        Object.entries(templates).forEach(([key, tpl]) => {
          const opt = document.createElement('option');
          opt.value = key;
          opt.textContent = tpl.label;
          tplSelect.appendChild(opt);
        });
        tplSelect.value = 'none';
        tplSelect.addEventListener('change', (e) => {
          npcTemplateKey = e.target.value || 'none';
        });
        tplField.appendChild(tplLabel);
        tplField.appendChild(tplSelect);
        content.appendChild(tplField);
      }

      const btnWrap = document.createElement('div');
      btnWrap.style.display = 'flex';
      btnWrap.style.flexDirection = 'column';
      btnWrap.style.gap = '8px';

      function applyAndClose(name) {
        const camp = currentCampaign();
        if (type === 'pc') {
          const parts = name.split(' ');
          const ent = newEntity(camp, 'pc');
          ent.firstName = parts[0] || name;
          ent.lastName = parts.slice(1).join(' ') || '';
          ent.name = name;
          // Track ownership when in player mode
          if (!state.gmMode) camp.playerOwnedPcId = ent.id;
          appendLog('Created PC', ent.id);
          saveAndRefresh();
          selectEntity(ent.id);
        } else if (type === 'npc') {
          const ent = newEntity(camp, 'npc');
          ent.name = name;
          ent.role = 'NPC';
          applyNpcTemplate(ent, npcTemplateKey);
          appendLog('Created NPC', ent.id);
          saveAndRefresh();
          selectEntity(ent.id);
        } else {
          const ent = newEntity(camp, 'org');
          ent.name = name;
          appendLog('Created organisation', ent.id);

          // Auto-generate one NPC member with a "Member Of" relationship
          const npc = newEntity(camp, 'npc');
          npc.name = `${randomFrom(DROW_FIRST_NAMES)} ${randomFrom(DROW_LAST_NAMES)}`;
          npc.role = 'Member';
          ent.members.push(npc.id);
          newRelationship(camp, npc.id, ent.id, 'Member Of');
          appendLog('Created NPC member', npc.id);

          saveAndRefresh();
          selectEntity(ent.id);
        }
        overlay.classList.add('hidden');
        modal.classList.add('hidden');
      }

      if (type !== 'org') {
        // Drow name button
        const drowBtn = document.createElement('button');
        drowBtn.className = 'modal-submit';
        drowBtn.style.marginBottom = '0';
        drowBtn.innerHTML = '<span style="opacity:0.6;margin-right:6px">⬡</span> Generate Drow Name';
        drowBtn.addEventListener('click', () => {
          const n = `${randomFrom(DROW_FIRST_NAMES)} ${randomFrom(DROW_LAST_NAMES)}`;
          applyAndClose(n);
        });
        btnWrap.appendChild(drowBtn);

        // Modern name button
        const modernBtn = document.createElement('button');
        modernBtn.className = 'modal-submit';
        modernBtn.style.marginBottom = '0';
        modernBtn.style.background = 'var(--spire-mid)';
        modernBtn.style.borderColor = 'var(--spire-border)';
        modernBtn.style.color = 'var(--spire-text)';
        modernBtn.innerHTML = '<span style="opacity:0.6;margin-right:6px">👤</span> Generate Modern Name';
        modernBtn.addEventListener('click', () => {
          const n = `${randomFrom(MODERN_FIRST_NAMES)} ${randomFrom(MODERN_LAST_NAMES)}`;
          applyAndClose(n);
        });
        btnWrap.appendChild(modernBtn);
      } else {
        const orgBtn = document.createElement('button');
        orgBtn.className = 'modal-submit';
        orgBtn.innerHTML = '<span style="opacity:0.6;margin-right:6px">⬡</span> Generate Random Name';
        orgBtn.addEventListener('click', () => applyAndClose(randomOrgName()));
        btnWrap.appendChild(orgBtn);
      }

      // Divider
      const divider = document.createElement('div');
      divider.style.display = 'flex';
      divider.style.alignItems = 'center';
      divider.style.gap = '8px';
      divider.style.margin = '4px 0';
      divider.innerHTML = '<hr style="flex:1;border-color:var(--spire-border)"><span style="color:var(--spire-muted);font-size:0.8rem">or enter manually</span><hr style="flex:1;border-color:var(--spire-border)">';
      btnWrap.appendChild(divider);

      // Manual name input
      const customInput = document.createElement('input');
      customInput.type = 'text';
      customInput.placeholder = type === 'org' ? 'Organisation name…' : 'Full name…';
      customInput.style.marginBottom = '0';
      btnWrap.appendChild(customInput);

      const customBtn = document.createElement('button');
      customBtn.className = 'modal-submit';
      customBtn.style.background = 'var(--spire-mid)';
      customBtn.style.borderColor = 'var(--spire-border)';
      customBtn.style.color = 'var(--spire-text)';
      customBtn.textContent = 'Use This Name';
      customBtn.addEventListener('click', () => {
        const val = customInput.value.trim();
        if (!val) { customInput.focus(); return; }
        applyAndClose(val);
      });
      customInput.addEventListener('keydown', e => {
        if (e.key === 'Enter') customBtn.click();
      });
      btnWrap.appendChild(customBtn);

      content.appendChild(btnWrap);
    }

    // Add entity buttons
    document.getElementById('add-pc-btn').addEventListener('click', () => openNamePickerModal('pc'));
    document.getElementById('add-npc-btn').addEventListener('click', () => openNamePickerModal('npc'));
    document.getElementById('add-org-btn').addEventListener('click', () => openNamePickerModal('org'));

    // Player-mode: claim an existing PC, open owned one, or create a new PC
    document.getElementById('player-add-pc-btn').addEventListener('click', () => {
      const camp = currentCampaign();
      const ownedId = camp.playerOwnedPcId;
      if (ownedId && camp.entities[ownedId]) {
        selectEntity(ownedId);
        return;
      }
      const pcs = Object.values(camp.entities).filter(e => e.type === 'pc');
      if (pcs.length) {
        openClaimPCModal(() => openNamePickerModal('pc'));
        return;
      }
      openNamePickerModal('pc');
    });
    document.getElementById('player-submit-npc-btn').addEventListener('click', () => {
      openPlayerSubmitNPCModal();
    });
    // Web filter pills
    document.querySelectorAll('#web-filter-bar .filter-pill[data-type]').forEach(pill => {
      pill.addEventListener('click', (e) => {
        e.preventDefault();
        const type = pill.dataset.type;
        webFilter[type] = !webFilter[type];
        syncWebFilterPills();
        updateGraph();
      });
    });
    syncWebFilterPills();
    const focusToggle = document.getElementById('filter-focus-toggle');
    if (focusToggle) {
      focusToggle.addEventListener('click', () => {
        setGraphFocusMode(!graphState.focusMode);
      });
    }

    // Dice roller
    document.getElementById('dice-btn').addEventListener('click', openDiceRoller);
    // Settings
    document.getElementById('settings-btn').addEventListener('click', openSettingsModal);
    // New session
    document.getElementById('new-session-btn').addEventListener('click', async () => {
      const ok = await askConfirm('Start a new session? Core abilities for all PCs will be reset.', 'New Session');
      if (!ok) return;
      startNewSession();
    });
    // Session note
    const noteInput = document.getElementById('log-note-input');
    const logSearchInput = document.getElementById('log-search-input');
    if (logSearchInput) {
      logSearchInput.addEventListener('input', (e) => {
        logFilterState.query = e.target.value || '';
        logFilterState.page = 1;
        renderLog();
      });
    }
    const logSessionFilter = document.getElementById('log-session-filter');
    if (logSessionFilter) {
      logSessionFilter.addEventListener('change', (e) => {
        logFilterState.session = e.target.value || 'all';
        logFilterState.page = 1;
        renderLog();
      });
    }
    const logActorFilter = document.getElementById('log-actor-filter');
    if (logActorFilter) {
      logActorFilter.addEventListener('change', (e) => {
        logFilterState.actor = e.target.value || 'all';
        logFilterState.page = 1;
        renderLog();
      });
    }
    const logTypeFilter = document.getElementById('log-type-filter');
    if (logTypeFilter) {
      logTypeFilter.addEventListener('change', (e) => {
        logFilterState.actionType = e.target.value || 'all';
        logFilterState.page = 1;
        renderLog();
      });
    }
    document.getElementById('log-note-btn').addEventListener('click', () => {
      const val = noteInput.value.trim();
      if (!val) return;
      appendLog(val, '', 'note');
      noteInput.value = '';
      saveCampaigns();
      renderLog();
    });
    noteInput.addEventListener('keydown', e => {
      if (e.key === 'Enter') document.getElementById('log-note-btn').click();
    });
    // GM toggle
    document.getElementById('gm-toggle').addEventListener('change', e => {
      toggleGMMode(e.target.checked);
    });
    // Dark mode toggle
    document.getElementById('dark-mode-btn').addEventListener('click', () => {
      toggleDarkMode();
    });
    // Save button (explicit save)
    document.getElementById('save-btn').addEventListener('click', () => {
      void attemptManualSave();
    });
    const saveRetryBtn = document.getElementById('save-retry-btn');
    if (saveRetryBtn) {
      saveRetryBtn.addEventListener('click', () => {
        void attemptManualSave();
      });
    }
    const syncReloadBtn = document.getElementById('sync-reload-btn');
    if (syncReloadBtn) {
      syncReloadBtn.addEventListener('click', () => {
        reloadCampaignFromStorage();
      });
    }
    const syncForceBtn = document.getElementById('sync-force-btn');
    if (syncForceBtn) {
      syncForceBtn.addEventListener('click', async () => {
        const ok = await askConfirm(
          'Force overwrite using data from this tab?',
          'Force Overwrite'
        );
        if (!ok) return;
        forceOverwriteSave();
      });
    }
    const syncIndicator = document.getElementById('sync-conflict-indicator');
    if (syncIndicator) {
      syncIndicator.addEventListener('click', () => {
        openSyncConflictModal();
      });
    }
    const syncQueueBtn = document.getElementById('sync-queue-btn');
    if (syncQueueBtn) {
      syncQueueBtn.addEventListener('click', () => {
        openSyncQueueModal();
      });
    }
    const undoBtn = document.getElementById('undo-btn');
    if (undoBtn) {
      undoBtn.addEventListener('click', () => {
        undoLastDestructiveAction();
      });
    }
    const redoBtn = document.getElementById('redo-btn');
    if (redoBtn) {
      redoBtn.addEventListener('click', () => {
        redoLastDestructiveAction();
      });
    }
    const shortcutsBtn = document.getElementById('shortcuts-btn');
    if (shortcutsBtn) {
      shortcutsBtn.addEventListener('click', () => openShortcutHelpModal());
    }
    // Node selection dropdown: jump to a node from the graph or sheet
    const nodeSel = document.getElementById('node-select');
    if (nodeSel) {
      nodeSel.addEventListener('change', e => {
        const id = e.target.value;
        if (!id) return;
        // Select the entity in the sidebar and sheet
        selectEntity(id);
        // Center the custom graph view on the selected node
        centerOnNode(id);
      });
    }
    const graphViewSel = document.getElementById('graph-view-select');
    if (graphViewSel) {
      graphViewSel.addEventListener('change', e => {
        const name = e.target.value;
        if (!name) return;
        applyGraphView(name);
      });
    }
    const saveGraphViewBtn = document.getElementById('save-graph-view-btn');
    if (saveGraphViewBtn) {
      saveGraphViewBtn.addEventListener('click', async () => {
        const current = document.getElementById('graph-view-select')?.value || '';
        const name = await askPrompt('Name this graph view preset:', current, { title: 'Save Graph View', submitText: 'Save' });
        if (!name || !name.trim()) return;
        saveCurrentGraphView(name.trim());
      });
    }
    const deleteGraphViewBtn = document.getElementById('delete-graph-view-btn');
    if (deleteGraphViewBtn) {
      deleteGraphViewBtn.addEventListener('click', async () => {
        const sel = document.getElementById('graph-view-select');
        const name = sel ? sel.value : '';
        if (!name) {
          showToast('Select a graph view to delete.', 'warn');
          return;
        }
        const ok = await askConfirm('Delete graph view "' + name + '"?', 'Delete Graph View');
        if (!ok) return;
        const camp = currentCampaign();
        ensureGraphViews(camp);
        delete camp.graphViews[name];
        saveCampaigns();
        populateGraphViewSelect();
        showToast('Deleted graph view "' + name + '".', 'info');
      });
    }
    // Close modal via the X button
    const closeModalBtn = document.getElementById('close-modal');
    if (closeModalBtn) {
      closeModalBtn.addEventListener('click', () => {
        closeModal();
      });
    }
    // Export button
    document.getElementById('export-btn').addEventListener('click', async () => {
      // Default is player-safe; GM must explicitly opt into full export.
      const gm = state.gmMode
        ? await askConfirm('Include GM-only and secret data? Cancel exports player-safe copy.', 'Export GM Data')
        : false;
      const json = exportCampaign(gm);
      const blob = new Blob([json], { type: 'application/json' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = `${currentCampaign().name.replace(/\s+/g, '_')}_${gm ? 'gm' : 'player'}.json`;
      link.click();
    });
    // Import button triggers file input
    document.getElementById('import-btn').addEventListener('click', () => {
      document.getElementById('import-input').click();
    });
    document.getElementById('import-input').addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (evt) => {
        importCampaign(evt.target.result);
        e.target.value = '';
      };
      reader.readAsText(file);
    });
    // Add relationship button: open relationship modal with no preset source
    document.getElementById('add-rel-btn').addEventListener('click', () => {
      // Pre-fill source from currently selected entity or selected canvas node
      const preSource = state.selectedEntityId || graphSelectedNodeId || null;
      openRelModal(preSource);
    });
    // Export PNG button
    document.getElementById('export-png-btn').addEventListener('click', async () => {
      const full = await askConfirm('Export full graph bounds? Cancel exports only current view.', 'Export Graph PNG');
      exportGraphPNG(full);
    });
    // Layout button: run a simple force-directed layout on the custom graph
    document.getElementById('layout-btn').addEventListener('click', () => {
      forceLayout();
    });
    // Clear log
    document.getElementById('clear-log-btn').addEventListener('click', async () => {
      const ok = await askConfirm('Clear activity log?', 'Clear Log');
      if (!ok) return;
      captureUndoSnapshot('Clear activity log');
      currentCampaign().logs = [];
      saveAndRefresh();
    });
    const recapBtn = document.getElementById('generate-recap-btn');
    if (recapBtn) {
      recapBtn.addEventListener('click', () => {
        openSessionRecapModal();
      });
    }
    const exportLogCsvBtn = document.getElementById('export-log-csv-btn');
    if (exportLogCsvBtn) {
      exportLogCsvBtn.addEventListener('click', () => {
        exportFilteredLogCsv();
      });
    }
    const sendMessageBtn = document.getElementById('send-message-btn');
    const messageInput = document.getElementById('message-input');
    const targetSelect = document.getElementById('message-target-select');
    if (sendMessageBtn && messageInput && targetSelect) {
      const sendMessage = () => {
        const text = (messageInput.value || '').trim();
        if (!text) return;
        appendMessage(targetSelect.value || 'party', text);
        messageInput.value = '';
        saveAndRefresh();
        renderMessages();
      };
      sendMessageBtn.addEventListener('click', sendMessage);
      messageInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') sendMessage();
      });
    }
    const messagesFilterSel = document.getElementById('messages-filter-select');
    if (messagesFilterSel) {
      messagesFilterSel.value = messageFilterState.mode || 'all';
      messagesFilterSel.addEventListener('change', (e) => {
        messageFilterState.mode = e.target.value || 'all';
        renderMessages();
      });
    }
    const entitySortSel = document.getElementById('entity-sort-select');
    if (entitySortSel) {
      entitySortSel.value = currentCampaign().entitySort || 'manual';
      entitySortSel.addEventListener('change', (e) => {
        const camp = currentCampaign();
        camp.entitySort = ['manual', 'name', 'pinned'].includes(e.target.value) ? e.target.value : 'manual';
        saveCampaigns();
        renderEntityLists();
      });
    }
    const entityPinOnly = document.getElementById('entity-pin-only');
    if (entityPinOnly) {
      entityPinOnly.checked = !!currentCampaign().entityPinnedOnly;
      entityPinOnly.addEventListener('change', (e) => {
        currentCampaign().entityPinnedOnly = !!e.target.checked;
        saveCampaigns();
        renderEntityLists();
      });
    }
    const markMessagesReadBtn = document.getElementById('mark-messages-read-btn');
    if (markMessagesReadBtn) {
      markMessagesReadBtn.addEventListener('click', () => {
        const camp = currentCampaign();
        if (!camp || !Array.isArray(camp.messages)) return;
        const visible = camp.messages.filter(canViewMessage).filter(messageMatchesFilter);
        const changed = markVisibleMessagesRead(visible);
        if (changed) saveCampaigns();
        renderMessages();
      });
    }
    const clearMessagesBtn = document.getElementById('clear-messages-btn');
    if (clearMessagesBtn) {
      clearMessagesBtn.addEventListener('click', async () => {
        const ok = await askConfirm('Clear all table messages?', 'Clear Messages');
        if (!ok) return;
        const camp = currentCampaign();
        captureUndoSnapshot('Clear table messages');
        camp.messages = [];
        saveAndRefresh();
      });
    }
    const gmAddSessionNoteBtn = document.getElementById('gm-add-session-note-btn');
    if (gmAddSessionNoteBtn) {
      gmAddSessionNoteBtn.addEventListener('click', () => {
        if (!state.gmMode) return;
        const camp = currentCampaign();
        ensureGmNotesStore(camp).push(createGmNote(camp, 'session', camp.currentSession || 1));
        appendLog('Added GM note', '');
        saveAndRefresh();
      });
    }
    const gmAddGlobalNoteBtn = document.getElementById('gm-add-global-note-btn');
    if (gmAddGlobalNoteBtn) {
      gmAddGlobalNoteBtn.addEventListener('click', () => {
        if (!state.gmMode) return;
        const camp = currentCampaign();
        ensureGmNotesStore(camp).push(createGmNote(camp, 'global'));
        appendLog('Added GM note', '');
        saveAndRefresh();
      });
    }
    // Close inspector
    document.getElementById('close-inspector').addEventListener('click', () => {
      state.selectedRelId = null;
      document.getElementById('inspector').classList.add('hidden');
    });
    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
      if (e.key === '/' && !e.target.matches('input, textarea')) {
        document.getElementById('search-input').focus();
        e.preventDefault();
      } else if ((e.key === 'ArrowDown' || e.key === 'ArrowUp') && !e.target.matches('input, textarea, select')) {
        const modalOpen = !document.getElementById('modal')?.classList.contains('hidden');
        if (modalOpen) return;
        selectAdjacentEntity(e.key === 'ArrowDown' ? 1 : -1);
        e.preventDefault();
      } else if ((e.key === 'ArrowLeft' || e.key === 'ArrowRight') && !e.target.matches('input, textarea, select')) {
        const modalOpen = !document.getElementById('modal')?.classList.contains('hidden');
        if (modalOpen) return;
        selectAdjacentTab(e.key === 'ArrowRight' ? 1 : -1);
        e.preventDefault();
      } else if (e.key.toLowerCase() === 't' && e.shiftKey && !e.target.matches('input, textarea, select')) {
        const modalOpen = !document.getElementById('modal')?.classList.contains('hidden');
        if (modalOpen) return;
        quickAddTaskToSelectedPC();
        e.preventDefault();
      } else if (e.key === '?' && !e.target.matches('input, textarea')) {
        openShortcutHelpModal();
        e.preventDefault();
      } else if (e.key === 'Escape') {
        // Close inspector or modals
        document.getElementById('inspector').classList.add('hidden');
      } else if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.key.toLowerCase() === 'z' && !e.target.matches('input, textarea, select')) {
        e.preventDefault();
        undoLastDestructiveAction();
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        void attemptManualSave();
      } else if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.key.toLowerCase() === 'y' && !e.target.matches('input, textarea, select')) {
        e.preventDefault();
        redoLastDestructiveAction();
      } else if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'z' && !e.target.matches('input, textarea, select')) {
        e.preventDefault();
        redoLastDestructiveAction();
      }
    });
    // Tab switching
    document.querySelectorAll('.tab-link').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.tab-link').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const target = btn.dataset.tab;
        document.querySelectorAll('.tab-page').forEach(page => {
          if (page.id === target) page.classList.add('active'); else page.classList.remove('active');
        });
        // When switching to graph view, resize the custom canvas and redraw
        if (target === 'web-view') {
          setTimeout(() => { resizeGraphCanvas(); }, 50);
          updateGraph();
          // Auto-dismiss relationship inspector when leaving web
        }
        if (target === 'log-view') {
          renderSessionPrep();
        }
        if (target === 'messages-view') {
          renderMessages();
        }
        if (target === 'gm-notes-view') {
          renderGMNotes();
        }
        // Auto-dismiss inspector when leaving web view
        if (target !== 'web-view') {
          document.getElementById('inspector').classList.add('hidden');
        }
        ensureAccessibilitySemantics();
      });
    });
  }

  /**
   * Choose a random element from an array.
   */
  function randomFrom(arr) {
    return arr[Math.floor(Math.random() * arr.length)];
  }

  /**
   * Perform initialisation once campaigns have been loaded or imported.
   */
  function initAfterLoad() {
    const camp = currentCampaign();
    state.syncConfig = loadSyncConfig();
    state.telemetryConfig = loadTelemetryConfig();
    if (!state.currentUser) {
      closeRealtimeChannel();
    } else {
      initRealtimeChannel();
    }
    if (logFilterState.actor === undefined) logFilterState.actor = 'all';
    if (logFilterState.actionType === undefined) logFilterState.actionType = 'all';
    setSyncConflictWarning(false);
    if (!Array.isArray(camp.clocks)) camp.clocks = [];
    if (camp.owner === undefined || camp.owner === null) camp.owner = state.currentUser || null;
    if (!Array.isArray(camp.gmUsers) || !camp.gmUsers.length) {
      camp.gmUsers = camp.owner ? [camp.owner] : (state.currentUser ? [state.currentUser] : []);
    }
    if (camp.owner && !camp.gmUsers.includes(camp.owner)) camp.gmUsers.unshift(camp.owner);
    if (!Array.isArray(camp.memberUsers)) camp.memberUsers = camp.owner ? [camp.owner] : [];
    if (camp.owner && !camp.memberUsers.includes(camp.owner)) camp.memberUsers.unshift(camp.owner);
    if (camp.sourceInviteCode) {
      const code = String(camp.sourceInviteCode).toUpperCase();
      const sharedRec = loadSharedInvites()[code];
      const sharedMembers = Array.isArray(sharedRec?.data?.memberUsers) ? sharedRec.data.memberUsers : [];
      if (sharedMembers.length && state.currentUser && !sharedMembers.includes(state.currentUser)) {
        showToast('Access revoked for this joined campaign.', 'warn');
        return;
      }
    }
    if (state.currentUser && camp.memberUsers.length && !camp.memberUsers.includes(state.currentUser)) {
      showToast('You no longer have access to this campaign.', 'warn');
      renderModeScreenCampaigns();
      return;
    }
    renderCampaignName();
    applyModeClasses();
    toggleDarkMode(camp.darkMode);
    renderEntityLists();
    renderSheetView();
    setupGraph();
    updateGraph();
    renderLog();
    renderMessages();
    renderGMNotes();
    renderSessionPrep();
    updatePendingBadge();
    updatePlayerPCButtonState();
    updateFocusToggleUI();
    // Update session badge
    const badge = document.getElementById('log-session-badge');
    if (badge) badge.textContent = 'Session ' + (camp.currentSession || 1);
    // Migrate missing fields on old campaigns
    if (!camp.currentSession) camp.currentSession = 1;
    if (camp.allowPlayerEditing === undefined) camp.allowPlayerEditing = true;
    if (!camp.gmPin) camp.gmPin = '';
    if (camp.ministryAttention === undefined) camp.ministryAttention = 0;
    if (!Array.isArray(camp.memberUsers)) camp.memberUsers = [];
    if (camp.owner && !camp.memberUsers.includes(camp.owner)) camp.memberUsers.unshift(camp.owner);
    if (camp.inviteCode === undefined) camp.inviteCode = '';
    if (camp.sourceInviteCode === undefined) camp.sourceInviteCode = '';
    if (camp.cloudCampaignId === undefined) camp.cloudCampaignId = '';
    if (!camp.rulesProfile) camp.rulesProfile = 'Core';
    if (!camp.customRules) {
      camp.customRules = {
        difficultyDowngrades: true,
        falloutCheckOnStress: true,
        clearStressOnFallout: true
      };
    }
    if (!camp.falloutGuidance) camp.falloutGuidance = defaultFalloutGuidance();
    if (!camp.graphViews) camp.graphViews = {};
    if (!camp.sectionCollapse) camp.sectionCollapse = {};
    if (!camp.taskViewByPc || typeof camp.taskViewByPc !== 'object') camp.taskViewByPc = {};
    if (!camp.taskSortByPc || typeof camp.taskSortByPc !== 'object') camp.taskSortByPc = {};
    if (camp.sessionPrepTasksOnly === undefined) camp.sessionPrepTasksOnly = false;
    if (!['manual', 'name', 'pinned'].includes(camp.entitySort)) camp.entitySort = 'manual';
    if (camp.entityPinnedOnly === undefined) camp.entityPinnedOnly = false;
    if (!Array.isArray(camp.messages)) camp.messages = [];
    if (!Array.isArray(camp.gmNotes)) camp.gmNotes = [];
    if (camp.lastScenePrompt === undefined) camp.lastScenePrompt = null;
    if (!Array.isArray(camp.scenarioPacks)) camp.scenarioPacks = [];
    if (!camp.uiTipsDismissed || typeof camp.uiTipsDismissed !== 'object') camp.uiTipsDismissed = {};
    if (!camp.relationshipUndo || typeof camp.relationshipUndo !== 'object') camp.relationshipUndo = {};
    if (!camp.relationshipRedo || typeof camp.relationshipRedo !== 'object') camp.relationshipRedo = {};
    if (!Array.isArray(camp.undoStack)) camp.undoStack = [];
    if (!Array.isArray(camp.redoStack)) camp.redoStack = [];
    updateMessagesUnreadBadge();
    updateUndoButtonState();
    ensureAutoGrowTextareas(document);
    ensureAccessibilityLabels(document);
    ensureAccessibilitySemantics();
    updateSyncQueueButton();
    if (navigator.onLine) flushPendingSaveOps();
  }

  /**
   * Apply body classes based on current GM/player mode.
   */
  function applyModeClasses() {
    document.body.classList.toggle('gm-mode', state.gmMode);
    document.body.classList.toggle('player-mode', !state.gmMode);
    // Update mode badge
    const badge = document.getElementById('mode-badge');
    if (badge) {
      badge.textContent = state.gmMode ? 'GM' : 'Player';
      badge.className = 'mode-badge ' + (state.gmMode ? 'mode-badge-gm' : 'mode-badge-player');
    }
    // Update top nav accent CSS vars
    document.documentElement.style.setProperty('--accent',
      state.gmMode ? 'var(--gm-accent)' : 'var(--pl-accent)');
    document.documentElement.style.setProperty('--accent-lo',
      state.gmMode ? 'var(--gm-accent-lo)' : 'var(--pl-accent-lo)');
    document.documentElement.style.setProperty('--accent-hi',
      state.gmMode ? 'var(--gm-accent-hi)' : 'var(--pl-accent-hi)');
    document.documentElement.style.setProperty('--accent-glow',
      state.gmMode ? 'var(--gm-glow)' : 'var(--pl-glow)');
    const filterBar = document.getElementById('web-filter-bar');
    const gmNotesTab = document.querySelector('.tab-link[data-tab="gm-notes-view"]');
    const gmNotesPage = document.getElementById('gm-notes-view');
    // Keep GM-only filter controls hidden for players, but allow players
    // to access the web tab itself (player-safe filtering already applies).
    if (!state.gmMode) {
      if (filterBar) filterBar.style.display = 'none';
      if (gmNotesTab) gmNotesTab.style.display = 'none';
      if (gmNotesPage && gmNotesPage.classList.contains('active')) {
        const sheetsTab = document.querySelector('.tab-link[data-tab="sheets-view"]');
        if (sheetsTab) sheetsTab.click();
      }
    } else {
      if (filterBar) filterBar.style.display = '';
      if (gmNotesTab) gmNotesTab.style.display = '';
    }
    updatePlayerPCButtonState();
    updateFocusToggleUI();
  }

  /* -----------------------------------------------
     AUTH SCREEN
  ----------------------------------------------- */

  function showAuthScreen() {
    const auth = document.getElementById('auth-screen');
    const mode = document.getElementById('mode-screen');
    const main = document.getElementById('main-app');
    if (auth) auth.style.display = 'flex';
    if (mode) mode.style.display = 'none';
    if (main) main.classList.add('hidden');
  }

  function setAuthMessage(message, isError = false) {
    const msg = document.getElementById('auth-message');
    if (!msg) return;
    msg.textContent = message || '';
    msg.classList.toggle('error', !!isError);
  }

  function setAuthMode(mode) {
    authMode = (mode === 'register') ? 'register' : 'login';
    const loginTab = document.getElementById('auth-tab-login');
    const regTab = document.getElementById('auth-tab-register');
    const confirmWrap = document.getElementById('auth-confirm-wrap');
    const submitBtn = document.getElementById('auth-submit-btn');
    const password = document.getElementById('auth-password');
    if (loginTab) loginTab.classList.toggle('active', authMode === 'login');
    if (regTab) regTab.classList.toggle('active', authMode === 'register');
    if (confirmWrap) confirmWrap.classList.toggle('hidden', authMode !== 'register');
    if (submitBtn) submitBtn.textContent = authMode === 'register' ? 'Create User' : 'Login';
    if (password) password.autocomplete = authMode === 'register' ? 'new-password' : 'current-password';
    setAuthMessage('');
  }

  function updateModeUserRow() {
    const label = document.getElementById('mode-current-user');
    if (label) label.textContent = state.currentUser ? `Signed in as ${state.currentUser}` : 'Not signed in';
  }

  async function handleAuthSubmit() {
    const userInput = document.getElementById('auth-username');
    const passInput = document.getElementById('auth-password');
    const confirmInput = document.getElementById('auth-password-confirm');
    const username = (userInput ? userInput.value : '').trim();
    const password = passInput ? passInput.value : '';
    const confirm = confirmInput ? confirmInput.value : '';

    if (!username) {
      setAuthMessage('Username is required.', true);
      return;
    }
    if (username.length < 3) {
      setAuthMessage('Username must be at least 3 characters.', true);
      return;
    }
    if (!password || password.length < 6) {
      setAuthMessage('Password must be at least 6 characters.', true);
      return;
    }

    if (canUseOnlineApi()) {
      if (!initOnlineClient()) {
        setAuthMessage('Cloud auth is configured incorrectly.', true);
        return;
      }
      try {
        const email = usernameToAuthEmail(username);
        if (authMode === 'register') {
          if (password !== confirm) {
            setAuthMessage('Passwords do not match.', true);
            return;
          }
          await window.SpireOnlineClient.signUp(email, password, username, 'gm');
          await window.SpireOnlineClient.signIn(email, password);
          const onlineUser = await window.SpireOnlineClient.currentUser();
          state.currentUser = usernameFromOnlineUser(onlineUser, username);
          saveUsers();
          setAuthMessage('Cloud account created.');
          await enterModeScreenForCurrentUser();
          return;
        }
        await window.SpireOnlineClient.signIn(email, password);
        const onlineUser = await window.SpireOnlineClient.currentUser();
        state.currentUser = usernameFromOnlineUser(onlineUser, username);
        saveUsers();
        setAuthMessage('');
        await enterModeScreenForCurrentUser();
        return;
      } catch (err) {
        setAuthMessage(err?.message || 'Cloud authentication failed.', true);
        return;
      }
    }

    if (authMode === 'register') {
      if (password !== confirm) {
        setAuthMessage('Passwords do not match.', true);
        return;
      }
      if (state.users[username]) {
        setAuthMessage('That username already exists.', true);
        return;
      }
      state.users[username] = {
        passwordHash: await hashPassword(password),
        createdAt: new Date().toISOString()
      };
      state.currentUser = username;
      saveUsers();
      setAuthMessage('Account created.');
      await enterModeScreenForCurrentUser();
      return;
    }

    const existing = state.users[username];
    if (!existing) {
      setAuthMessage('Invalid username or password.', true);
      return;
    }
    const enteredHash = await hashPassword(password);
    if (enteredHash !== existing.passwordHash) {
      setAuthMessage('Invalid username or password.', true);
      return;
    }
    state.currentUser = username;
    saveUsers();
    setAuthMessage('');
    await enterModeScreenForCurrentUser();
  }

  function setupAuthScreen() {
    const loginTab = document.getElementById('auth-tab-login');
    const regTab = document.getElementById('auth-tab-register');
    const submitBtn = document.getElementById('auth-submit-btn');
    const inputs = ['auth-username', 'auth-password', 'auth-password-confirm']
      .map(id => document.getElementById(id))
      .filter(Boolean);

    if (loginTab) loginTab.addEventListener('click', () => setAuthMode('login'));
    if (regTab) regTab.addEventListener('click', () => setAuthMode('register'));
    if (submitBtn) submitBtn.addEventListener('click', () => { void handleAuthSubmit(); });
    inputs.forEach((input) => {
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') void handleAuthSubmit();
      });
    });
    setAuthMode('login');
  }

  async function logoutCurrentUser() {
    saveCampaigns();
    if (canUseOnlineApi() && initOnlineClient()) {
      try {
        await window.SpireOnlineClient.signOut();
      } catch (_) {}
    }
    closeRealtimeChannel();
    state.currentUser = null;
    state.syncConfig = defaultSyncConfig();
    state.telemetryConfig = defaultTelemetryConfig();
    state.campaigns = {};
    state.currentCampaignId = null;
    saveUsers();
    const userInput = document.getElementById('auth-username');
    const passInput = document.getElementById('auth-password');
    const confirmInput = document.getElementById('auth-password-confirm');
    if (userInput) userInput.value = '';
    if (passInput) passInput.value = '';
    if (confirmInput) confirmInput.value = '';
    setAuthMode('login');
    showAuthScreen();
  }

  async function enterModeScreenForCurrentUser() {
    if (state.currentUser) {
      state.syncConfig = loadSyncConfig();
      state.telemetryConfig = loadTelemetryConfig();
      initRealtimeChannel();
    }
    const loadedCloud = await syncCampaignsFromCloud().catch(() => false);
    if (!loadedCloud) loadCampaigns();
    renderModeScreenCampaigns();
    updateModeUserRow();
    const camp = currentCampaign();
    if (camp) toggleDarkMode(!!camp.darkMode);
    const storageNote = document.getElementById('mode-storage-note');
    if (storageNote) {
      storageNote.textContent = canUseOnlineApi()
        ? 'Data sync: Supabase cloud + local cache'
        : 'Data saved locally in your browser';
    }
    const auth = document.getElementById('auth-screen');
    const mode = document.getElementById('mode-screen');
    const main = document.getElementById('main-app');
    if (auth) auth.style.display = 'none';
    if (mode) mode.style.display = 'flex';
    if (main) main.classList.add('hidden');
  }

  /* -----------------------------------------------
     MODE SELECTION SCREEN
  ----------------------------------------------- */

  /**
   * Populate the campaign selector on the mode screen.
   */
  function renderModeScreenCampaigns() {
    const sel = document.getElementById('mode-campaign-select');
    if (!sel) return;
    sel.innerHTML = '';
    Object.values(state.campaigns).forEach(camp => {
      const opt = document.createElement('option');
      opt.value = camp.id;
      const pcs = Object.values(camp.entities || {}).filter(e => e.type === 'pc').length;
      const npcs = Object.values(camp.entities || {}).filter(e => e.type === 'npc').length;
      const orgs = Object.values(camp.entities || {}).filter(e => e.type === 'org').length;
      const sess = camp.currentSession || 1;
      opt.textContent = `${camp.name}  [S${sess} · ${pcs}PC ${npcs}NPC ${orgs}Org]`;
      if (camp.id === state.currentCampaignId) opt.selected = true;
      sel.appendChild(opt);
    });
  }

  /**
   * Enter the main app in a given mode (GM or player).
   */
  async function enterApp(gmMode) {
    if (!state.currentUser) {
      showToast('Please login first.', 'warn');
      showAuthScreen();
      return;
    }
    const sel = document.getElementById('mode-campaign-select');
    if (sel && sel.value && state.campaigns[sel.value]) {
      state.currentCampaignId = sel.value;
    }
    const camp = currentCampaign();
    if (camp.owner === undefined || camp.owner === null) camp.owner = state.currentUser || null;
    if (!Array.isArray(camp.gmUsers) || !camp.gmUsers.length) {
      camp.gmUsers = camp.owner ? [camp.owner] : (state.currentUser ? [state.currentUser] : []);
    }
    if (camp.owner && !camp.gmUsers.includes(camp.owner)) camp.gmUsers.unshift(camp.owner);
    // Campaign-level GM authorization: only owner/authorized GM users.
    if (gmMode) {
      const gmUsers = Array.isArray(camp.gmUsers) ? camp.gmUsers : [];
      const isAuthorized = state.currentUser && gmUsers.includes(state.currentUser);
      if (!isAuthorized) {
        showToast('You are not authorized for GM access in this campaign. Entering as Player.', 'warn');
        gmMode = false;
      }
    }
    // GM PIN check
    if (gmMode && camp.gmPin && camp.gmPin.trim()) {
      const entered = await askPrompt('Enter GM PIN:', '', { title: 'GM Access', type: 'password', submitText: 'Enter' });
      if (entered !== camp.gmPin.trim()) {
        showToast('Incorrect PIN. Entering as Player.', 'warn');
        gmMode = false;
      }
    }
    state.gmMode = gmMode;
    camp.gmMode = gmMode;
    saveCampaigns();
    const auth = document.getElementById('auth-screen');
    if (auth) auth.style.display = 'none';
    document.getElementById('mode-screen').style.display = 'none';
    document.getElementById('main-app').classList.remove('hidden');
    initAfterLoad();
    setupEventListeners();
    setTimeout(resizeGraphCanvas, 100);
  }

  /**
   * Return to the mode selection screen.
   */
  function returnToModeScreen() {
    if (!state.currentUser) {
      showAuthScreen();
      return;
    }
    saveCampaigns();
    const auth = document.getElementById('auth-screen');
    if (auth) auth.style.display = 'none';
    document.getElementById('mode-screen').style.display = 'flex';
    document.getElementById('main-app').classList.add('hidden');
    updateModeUserRow();
    renderModeScreenCampaigns();
  }

  async function joinCampaignByInviteCode() {
    if (!state.currentUser) {
      showToast('Please login first.', 'warn');
      return;
    }
    const codeRaw = await askPrompt('Enter invite code:', '', { title: 'Join Campaign', submitText: 'Join' });
    const code = String(codeRaw || '').trim().toUpperCase();
    if (!code) return;
    if (canUseOnlineApi() && initOnlineClient()) {
      try {
        await window.SpireOnlineClient.joinCampaignWithCode(code);
        await syncCampaignsFromCloud();
        saveCampaigns({ force: true });
        renderModeScreenCampaigns();
        showToast('Joined cloud campaign via invite code.', 'info');
        return;
      } catch (err) {
        showToast(err?.message || 'Could not join campaign with that code.', 'warn');
        return;
      }
    }
    const shared = loadSharedInvites();
    const record = shared[code];
    if (!record || !record.data || !record.data.entities || !record.data.relationships) {
      showToast('Invite code not found.', 'warn');
      return;
    }
    const existing = Object.values(state.campaigns).find((c) => String(c.sourceInviteCode || '').toUpperCase() === code);
    let targetId = existing ? existing.id : generateId('camp');
    if (existing) {
      const replace = await askConfirm(
        `You already joined "${existing.name}" with this code. Replace local copy with latest shared version?`,
        'Update Joined Campaign'
      );
      if (!replace) targetId = generateId('camp');
    }

    const data = JSON.parse(JSON.stringify(record.data));
    data.id = targetId;
    data.sourceInviteCode = code;
    if (!Array.isArray(data.memberUsers)) data.memberUsers = [];
    if (!data.memberUsers.includes(state.currentUser)) data.memberUsers.push(state.currentUser);
    if (!Array.isArray(data.gmUsers)) data.gmUsers = [];
    if (!Array.isArray(data.undoStack)) data.undoStack = [];
    if (!Array.isArray(data.redoStack)) data.redoStack = [];
    state.campaigns[targetId] = data;
    state.currentCampaignId = targetId;

    // Update membership in shared record for owner visibility.
    if (!Array.isArray(record.data.memberUsers)) record.data.memberUsers = [];
    if (!record.data.memberUsers.includes(state.currentUser)) {
      record.data.memberUsers.push(state.currentUser);
      record.updatedAt = new Date().toISOString();
      shared[code] = record;
      saveSharedInvites(shared);
    }

    saveCampaigns();
    renderModeScreenCampaigns();
    showToast('Joined campaign via invite code.', 'info');
  }

  /**
   * Set up mode screen event listeners.
   */
  function setupModeScreen() {
    document.getElementById('enter-gm').addEventListener('click', () => { void enterApp(true); });
    document.getElementById('enter-player').addEventListener('click', () => { void enterApp(false); });

    document.getElementById('back-to-mode').addEventListener('click', returnToModeScreen);
    const logoutBtn = document.getElementById('mode-logout-user');
    if (logoutBtn) {
      logoutBtn.addEventListener('click', () => {
        void logoutCurrentUser();
      });
    }

    // Campaign rename on title click
    const titleEl = document.getElementById('campaign-name');
    if (titleEl) {
      titleEl.addEventListener('click', async () => {
        const camp = currentCampaign();
        const newName = await askPrompt('Campaign name:', camp.name, { title: 'Rename Campaign', submitText: 'Rename' });
        if (newName && newName.trim()) {
          camp.name = newName.trim();
          renderCampaignName();
          saveCampaigns();
          renderModeScreenCampaigns();
        }
      });
    }

    // New campaign button on mode screen
    document.getElementById('mode-new-campaign').addEventListener('click', async () => {
      const name = await askPrompt('New campaign name:', '', { title: 'New Campaign', submitText: 'Create' });
      if (!name || !name.trim()) return;
      if (canUseOnlineApi() && initOnlineClient()) {
        try {
          const cloudId = await window.SpireOnlineClient.createCampaign(name.trim());
          const camp = createCampaign(name.trim());
          camp.id = String(cloudId || camp.id);
          camp.cloudCampaignId = camp.id;
          state.campaigns[camp.id] = camp;
          state.currentCampaignId = camp.id;
          await persistCurrentCampaignToCloud();
          saveCampaigns();
          renderModeScreenCampaigns();
          showToast('Cloud campaign created.', 'info');
          return;
        } catch (err) {
          showToast(err?.message || 'Could not create cloud campaign.', 'warn');
          return;
        }
      }
      const camp = createCampaign(name.trim());
      state.campaigns[camp.id] = camp;
      state.currentCampaignId = camp.id;
      saveCampaigns();
      renderModeScreenCampaigns();
    });
    const joinBtn = document.getElementById('mode-join-code');
    if (joinBtn) {
      joinBtn.addEventListener('click', () => { void joinCampaignByInviteCode(); });
    }

    // Import on mode screen
    document.getElementById('mode-import-btn').addEventListener('click', () => {
      document.getElementById('mode-import-input').click();
    });
    document.getElementById('mode-delete-campaign').addEventListener('click', async () => {
      const camp = currentCampaign();
      if (!camp) return;
      const ok = await askConfirm('Delete "' + camp.name + '"? This cannot be undone.', 'Delete Campaign');
      if (!ok) return;
      delete state.campaigns[camp.id];
      const remaining = Object.keys(state.campaigns);
      if (remaining.length === 0) {
        const fresh = createCampaign('New Campaign');
        state.campaigns[fresh.id] = fresh;
        state.currentCampaignId = fresh.id;
      } else {
        state.currentCampaignId = remaining[0];
      }
      saveCampaigns();
      renderModeScreenCampaigns();
    });
    document.getElementById('mode-import-input').addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const data = JSON.parse(evt.target.result);
          if (!data.entities || !data.relationships) { showToast('Invalid file', 'warn'); return; }
          const newId = generateId('camp');
          data.id = newId;
          data.name = data.name || 'Imported Campaign';
          data.owner = state.currentUser || null;
          data.gmUsers = state.currentUser ? [state.currentUser] : [];
          data.memberUsers = state.currentUser ? [state.currentUser] : [];
          data.inviteCode = '';
          data.sourceInviteCode = '';
          if (!Array.isArray(data.undoStack)) data.undoStack = [];
          if (!Array.isArray(data.redoStack)) data.redoStack = [];
          if (!Array.isArray(data.gmNotes)) data.gmNotes = [];
          state.campaigns[newId] = data;
          state.currentCampaignId = newId;
          saveCampaigns();
          renderModeScreenCampaigns();
        } catch(err) { showToast('Import failed: ' + err.message, 'warn'); }
        e.target.value = '';
      };
      reader.readAsText(file);
    });
  }

  /**
   * Entry point.
   */
  async function init() {
    bindGlobalErrorHandlers();
    loadUsers();
    setupAuthScreen();
    setupModeScreen();
    if (!state.currentUser && canUseOnlineApi() && initOnlineClient()) {
      try {
        const onlineUser = await window.SpireOnlineClient.currentUser();
        if (onlineUser) {
          state.currentUser = usernameFromOnlineUser(onlineUser, '');
          saveUsers();
        }
      } catch (_) {
        // keep auth screen visible when no cloud session is available
      }
    }
    if (state.currentUser) {
      await enterModeScreenForCurrentUser();
    } else {
      showAuthScreen();
    }
  }

  // Start application when DOM is ready
  document.addEventListener('DOMContentLoaded', () => { void init(); });
})();
