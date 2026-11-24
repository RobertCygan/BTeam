/*
 * Kjønn ut fra fornavn — enkel ordliste over vanlige fornavn (norske og internasjonale).
 *
 * Brukes bare når Excel-filen ikke har en egen kolonne «Kjønn».
 * Navn skrives med små bokstaver og uten æ/ø/å (æ→ae, ø→o, å→a).
 * Legg gjerne til flere navn her.
 */
(function (global) {
  'use strict';

  const KVINNE = `
    aase agnes agnete aina aino alma amalie amina anastasia anastasiia ane anette anita ann anna anne
    annette anu asa asta astrid aud aurora ayan aisha bente berit birgit birgitte bjorg borghild brit
    brita britt camilla caroline cathrine cecilie dagny daria dariia dorthe edel eli elin elisabeth
    ella ellen else emilie emma eva fadumo fatima frida galina gerd gina gita grete greta gro gunhild
    gunn gunnhild guro halima halyna hanna hannah hawa hege heidi helen helena helene hilde hildur
    ida inga ingebjorg ingeborg ingelin inger ingrid ingunn inna irene irina iryna jenny johanne
    jorunn judith julia julie kamilla kari karin karina karoline katarina katariina kateryna katrine
    khadija kirsten kjersti kristin kristina kristine laila larisa larysa laxmi layla leila lena lene
    liisa lilja linda line linn lisa lise liudmila liudmyla liv ludmila lyudmila magnhild maja malin
    margarita margit margrete marharyta maria mariam marianne marie marina marit marte martha martine
    mary maryam maryna maya mebrat merete mette minna mona monica monika nadia nadiya nadja natalia
    nataliia natalya nina nora oddny oddrun oksana olena oleana olga olha paivi rahel ragnhild randi
    rebecca renate rigmor rita ruta ruth saba sahra sara sarah sari satu selam senait signe silje
    siri sissel siv sofia sofie solveig sophie stine sunniva svetlana svitlana synnove tatiana
    tetiana therese thea tiina tina tone tonje torill tove trine trude tsega turid unni valentyna
    vera veronika victoria vigdis viktoria viktoriia vivian wenche wenke yasmin yordanos yulia yuliia
    zahra
  `;

  const MANN = `
    abdi abdiaziz abdifatah abdikadir abdirahim abdirahman abdul abdullah ahmad ahmed akram
    aleksander alexander alf ali anatolii anders andreas andrei andrey andrii andriy arild arne arnt
    artem asbjorn bahadur berhane bishnu bjarne bjorn bohdan christian christoffer dagaan daniel
    david dawit denis denys dmitri dmytro egil eirik einar emil erik espen evgenii evgeny farah finn
    frode fredrik gedi geir george gunnar haakon hakon hans harald hari hassan havard helge henrik
    henry hussein ibrahim ievgenii igor ihor inge ismail ivan ivegenii jakob jacob jama james jan
    johan johann johannes john jon jonas jose josef joseph jostein juan juha jukka kare kenneth
    khalid kjell knut kristian kristoffer lars leif liban magnus mahad marcus marius mark markus
    martin matti mathias matthias maxim maksym mehari michael mikael mikko mohamed mohammad mohammed
    morten muhammad mustafa mykhailo mykola nikolai nils odd ola olav olaf ole oleg oleh oleksandr
    oleksii oleksiy omar osman oyvind oystein paal pal paul pavel pavlo pekka per peter petro petter
    rabi rabie ragnar ram rashid reidar richard robert roar rolf roman rune ruslan salim samuel
    sander sebastian serhii sergei sergey shabbir sigmund sigurd simon sindre snorre solomon steffen
    stein steinar sten stian stig subhi svein sverre taras terje tesfay tesfaye thomas thor timo
    tobias tom tommy tor torbjorn tore torgeir trond trygve ulf valter vasyl vegard vidar viktor
    victor vitalii vitaly vladimir volodymyr walter william yam yevgeniy yevhenii yohannes yonas
    yurii yuri yusuf
  `;

  // Ord som tyder på en tjeneste/organisasjon i stedet for en person.
  const ORGANISASJON = `
    nav helse omsorg tjeneste flyktningtjeneste flyktingstjeneste kommune kommunal barnevern
    boligkontor as sa stiftelse
  `;

  const toSet = (s) => new Set(s.split(/\s+/).filter(Boolean));
  const kvinne = toSet(KVINNE);
  const mann = toSet(MANN);
  const org = toSet(ORGANISASJON);

  /**
   * Gjetter kjønn ut fra et fullt navn (rekkefølgen på for- og etternavn spiller ingen rolle).
   * `fold` gjør om teksten til små bokstaver uten diakritiske tegn.
   * Returnerer 'K' (kvinne), 'M' (mann), 'org' (organisasjon) eller null (ukjent).
   */
  function gjett(navn, fold) {
    const ord = fold(navn).split(/[^a-z]+/).filter((w) => w.length > 1);
    if (ord.some((w) => org.has(w))) return 'org';
    let k = 0, m = 0;
    for (const w of ord) {
      if (kvinne.has(w)) k++;
      if (mann.has(w)) m++;
    }
    if (k > m) return 'K';
    if (m > k) return 'M';
    return null;
  }

  global.NavnKjonn = { gjett };
})(window);
