import * as XLSX from 'xlsx';

export interface DriveCodeEntry {
  name: string;
  driveCode: string;
  jdeCode: string;
  aliases?: string[];
  shortName?: string;
}

/**
 * Master catalog transcribed from the user's reference sheet (اسم الصنف | كود الصنف دريف | كود الصنف JDE).
 */
export const MASTER_DRIVE_CODES_CATALOG: DriveCodeEntry[] = [
  { name: 'أذرة صب - Yellow Corn 7.5%', shortName: 'أذرة صب', driveCode: '1', jdeCode: '11342', aliases: ['أذرة صب', 'اذرة صب', 'corn', 'yellow corn', 'yellow corn 7.5%', 'corn 7.5%'] },
  { name: 'اذرة صفراء محلى - Yellow Local Corn', shortName: 'اذرة صفراء محلى', driveCode: '42', jdeCode: '1186134', aliases: ['اذرة صفراء محلى', 'yellow local corn'] },
  { name: 'أذره اوكرانى - Ukrainian Corn', shortName: 'أذره اوكرانى', driveCode: '30', jdeCode: '11342', aliases: ['أذره اوكرانى', 'اذره اوكراني', 'ukrainian corn'] },
  { name: 'اذره سورجم (عويجة) - Sorghum bicolor Corn', shortName: 'اذره سورجم (عويجة)', driveCode: '18', jdeCode: '11359', aliases: ['اذره سورجم', 'عويجة', 'sorghum', 'sorghum bicolor corn'] },
  { name: 'صويا النور 46% - Soya', shortName: 'صويا النور 46%', driveCode: '34', jdeCode: '', aliases: ['صويا النور 46%', 'صويا النور'] },
  { name: 'صويا 46% - Soya', shortName: 'صويا 46 %', driveCode: '2', jdeCode: '913686', aliases: ['صويا 46%', 'صويا 46 %', 'local soya 46%', 'soya 46%', 'soybean meal 46%'] },
  { name: 'صويا 44% - Soya', shortName: 'صويا 44 %', driveCode: '3', jdeCode: '11345', aliases: ['صويا 44%', 'soya 44%'] },
  { name: 'فولفات - Soya full fat', shortName: 'صويا full fat', driveCode: '4', jdeCode: '875561', aliases: ['صويا full fat', 'فولفات', 'soya full fat', 'full fat soya'] },
  { name: 'صويا هاى فات - HI FAT', shortName: 'صويا هاى فات', driveCode: '44', jdeCode: '11346', aliases: ['صويا هاي فات', 'hi fat'] },
  { name: 'كسب دوار الشمس - Sunflower', shortName: 'كسب دوار الشمس', driveCode: '13', jdeCode: '11343', aliases: ['كسب دوار الشمس', 'sunflower'] },
  { name: 'حجر جيري - Limestone', shortName: 'حجر جيري', driveCode: '9', jdeCode: '750451', aliases: ['حجر جيري', 'limestone'] },
  { name: 'زيت صويا - Soya Oil', shortName: 'زيت صويا', driveCode: '5', jdeCode: '13100', aliases: ['زيت صويا', 'soya oil'] },
  { name: 'زيت سمك - Fish Oil', shortName: 'زيت سمك', driveCode: '24', jdeCode: '913758', aliases: ['زيت سمك', 'fish oil'] },
  { name: 'احماض دهنية Poultry Fat', shortName: 'احماض دهنية Poultry Fat', driveCode: '43', jdeCode: '755411', aliases: ['احماض دهنية', 'poultry fat'] },
  { name: 'زيت كتان Linseed oil', shortName: 'زيت كتان', driveCode: '47', jdeCode: '913740', aliases: ['زيت كتان', 'linseed oil'] },
  { name: 'مونوكالسيوم - Mono Calcium', shortName: 'mono calcium', driveCode: '8', jdeCode: '808819', aliases: ['مونوكالسيوم', 'مونو كالسيوم', 'mono calcium', 'monocalcium', 'mono'] },
  { name: 'مونوكالسيوم تركى - Turkish Mono Calcium', shortName: 'Turkish mono', driveCode: '17', jdeCode: '', aliases: ['مونوكالسيوم تركى', 'مونو كالسيوم تركي', 'turkish mono'] },
  { name: 'ثنائى تونسى - Dicalcium Phos', shortName: 'ثنائى تونسى', driveCode: '7', jdeCode: '11353', aliases: ['ثنائي تونسي', 'dicalcium phos'] },
  { name: 'ثنائى كالسيوم تركى - Dicalcium Phos', shortName: 'ثنائى كالسيوم تركى', driveCode: '35', jdeCode: '11354', aliases: ['ثنائي كالسيوم تركي'] },
  { name: 'ردة (نخالة) - Bran 14%', shortName: 'ردة (نخالة)', driveCode: '11', jdeCode: '11361', aliases: ['ردة', 'نخالة', 'bran 14%', 'wheat bran'] },
  { name: 'خليط زوائد سن - Wheat Screenings', shortName: 'خليط زوائد سن ( سن أبيض )', driveCode: '29', jdeCode: '889532', aliases: ['خليط زوائد سن', 'خليط زوائد سن ( سن أبيض )', 'سن أبيض', 'wheat screenings'] },
  { name: 'رجيع الكون 14% - Rice Bran', shortName: 'رجيع الكون', driveCode: '15', jdeCode: '841686', aliases: ['رجيع الكون', 'رجيع الكون 14%', 'rice bran', 'rice bran 14%'] },
  { name: 'جيليوتين معبأ', shortName: 'جيليوتين معبأ', driveCode: '48', jdeCode: '11349', aliases: ['جيليوتين معبأ', 'جلوتين', 'gluten'] },
  { name: 'D.D.G.S', shortName: 'DDG', driveCode: '12', jdeCode: '11350', aliases: ['ddgs', 'ddg', 'd.d.g.s'] },
  { name: 'بروتين معصور - Poultry Meal', shortName: 'بروتين معصور', driveCode: '19', jdeCode: '752042', aliases: ['بروتين معصور', 'poultry meal'] },
  { name: 'بروتين معصور ميكس', shortName: 'بروتين معصور ميكس', driveCode: '53', jdeCode: '752042', aliases: ['بروتين معصور ميكس'] },
  { name: 'مسحوق بروتين غامق Poultry Meal', shortName: 'مسحوق بروتين غامق', driveCode: '51', jdeCode: '1199875', aliases: ['مسحوق بروتين غامق'] },
  { name: 'مسحوق سمك', shortName: 'مسحوق سمك', driveCode: '22', jdeCode: '11352', aliases: ['مسحوق سمك', 'fish meal'] },
  { name: 'مسحوق سمك 65%', shortName: 'مسحوق سمك 65%', driveCode: '60', jdeCode: '1340040', aliases: ['مسحوق سمك 65%', 'fish meal 65%'] },
  { name: 'محسن شهية للحيوانات الاليفة Digest', shortName: 'محسن شهية للحيوانات الاليفة', driveCode: '56', jdeCode: '1172234', aliases: ['محسن شهية للحيوانات الاليفة', 'digest'] },
  { name: 'Rejected', shortName: 'Rejected', driveCode: '1001', jdeCode: '12385', aliases: ['rejected'] },
  { name: 'Rejected Aqua', shortName: 'Rejected Aqua', driveCode: '1002', jdeCode: '951025', aliases: ['rejected aqua', 'rejected aqua feed'] },
  { name: 'Rejected PETS', shortName: 'Rejected Pets', driveCode: '1003', jdeCode: '1147531', aliases: ['rejected pets'] },
  { name: 'الماء - Water', driveCode: '5050', jdeCode: '5050', aliases: ['الماء', 'water'] },
  { name: 'مخلفات زيت صويا - Soya Oil Waste', driveCode: '36', jdeCode: '1535306', aliases: ['مخلفات زيت صويا', 'soya oil waste'] },
  { name: 'Potato Powder', driveCode: '61', jdeCode: '' },
  { name: 'Digest محسن شهية للحيوانات الاليفة (petology)', driveCode: '703', jdeCode: '703' },
  { name: 'مخلفات خامات (رجيع)', driveCode: '39', jdeCode: '1118001' },
  { name: 'ثنائى صينى', driveCode: '49', jdeCode: '1188826' },
  { name: 'سرس 5%', driveCode: '54', jdeCode: '878201' },
  { name: 'كسر ارز', driveCode: '55', jdeCode: '787641', aliases: ['broken rice', 'كسر ارز'] },
  { name: 'Rejected Duck', driveCode: '1004', jdeCode: '1304751' },
  { name: 'Customer addition Rejected cat food', driveCode: '1005', jdeCode: '1005' },
  { name: 'مسحوق بروتين فاتح Poultry Meal', driveCode: '57', jdeCode: '1333350' },
  { name: 'اذرة بيضاء محلى - White Local Corn', driveCode: '41', jdeCode: '1133884' },
  { name: 'potato powder', driveCode: '58', jdeCode: '58' },
  { name: 'SWeet potato powder', driveCode: '59', jdeCode: '59' },
  { name: 'بسلة خضراء جافة', driveCode: '704', jdeCode: '704' },
  { name: 'بسلة صفراء جافة', driveCode: '705', jdeCode: '705' },
  { name: 'كسب بنجر Beat Pulp Dried', driveCode: '1006', jdeCode: '787659' },
  { name: 'مسحوق جزر مجفف', driveCode: '706', jdeCode: '706' },
  { name: 'زيت عباد شمس', driveCode: '52', jdeCode: '5296' },
  { name: 'محسن شهية للحيوانات الاليفة Digest خاص بالكلاب', driveCode: '70', jdeCode: '70' },
  { name: 'محسن شهية للحيوانات الاليفة Digest خاص بالقطط', driveCode: '71', jdeCode: '71' },
  { name: 'تنك بلاستيك 1000 لتر', driveCode: '2000', jdeCode: '15976' },
  { name: 'الاضافات', driveCode: '10101010', jdeCode: '10101010' },
  { name: 'فيتامينات بادي - VitaminStarter DSM', shortName: 'فيتامينات بادي DSM', driveCode: '313', jdeCode: '913504', aliases: ['فيتامينات بادي dsm', 'vitamin starter dsm'] },
  { name: 'فيتامينات نامى - Vitamin Grower DSM', driveCode: '314', jdeCode: '913483' },
  { name: 'فيتامينات ناهى - Vitamin Finisher DSM', driveCode: '315', jdeCode: '913491' },
  { name: 'فيتامينات تسمين بادي 1 كجم DSM', driveCode: '316', jdeCode: '1486647' },
  { name: 'فيتامينات نامى محلى - Vitamin Grower Local', driveCode: '372', jdeCode: '807146' },
  { name: 'فيتامينات ناهى محلى - Vitamin Finisher Local', driveCode: '572', jdeCode: '868310' },
  { name: 'كولين بودرة 60% - Choline Chloride', shortName: 'كولين بودرة 60%', driveCode: '308', jdeCode: '13559', aliases: ['كولين بودرة 60%', 'choline chloride 60%', 'choline 60%'] },
  { name: 'كولين بودرة 70%', driveCode: '309', jdeCode: '13560' },
  { name: 'Stay C 35', driveCode: '557', jdeCode: '13584' },
  { name: 'فيتامينات E 50', driveCode: '558', jdeCode: '13606' },
  { name: 'فيتامينات بادي 2 - VitaminStarter DSM', driveCode: '320', jdeCode: '' },
  { name: 'poultry star me', shortName: 'Poultry star', driveCode: '326', jdeCode: '13854', aliases: ['poultry star'] },
  { name: 'REMINERAL COBB BROILER STARTER', driveCode: '612', jdeCode: '1478655' },
  { name: 'املاح تسمين - Premix Mineral Broiler DSM', driveCode: '310', jdeCode: '913512' },
  { name: 'أملاح تسمين محلى - Premix Mineral Broiler Local', driveCode: '610', jdeCode: '1106341' },
  { name: 'Rovimix Broiler arbor acres', driveCode: '611', jdeCode: '1478663' },
  { name: 'أملاح تسمين إضافة 1.5 كجم DSM', driveCode: '700', jdeCode: '1486903', aliases: ['mineral premix broiler 1.5 kg'] },
  { name: 'ام- موبيلايز (Customer addition-Qantra) M-Mobilize', driveCode: '699', jdeCode: '699' },
  { name: 'شامب نيل', driveCode: '698', jdeCode: '1310481' },
  { name: 'ايدافكس بلس', driveCode: '697', jdeCode: '1310490' },
  { name: 'ميكوفيكس سليكت - Mycofix Select', shortName: 'ميكوفكس سليكت', driveCode: '307', jdeCode: '13845', aliases: ['ميكوفكس سليكت', 'mycofix select'] },
  { name: 'ميكوفكس بلس - Mycofix Plus', shortName: 'ميكوفكس بلس', driveCode: '306', jdeCode: '13844', aliases: ['ميكوفكس بلس', 'mycofix plus'] },
  { name: 'ميكوفكس سيكيور - MyCofix Secuer', shortName: 'ميكوفكس سيكيور', driveCode: '305', jdeCode: '855068', aliases: ['ميكوفكس سيكيور', 'mycofix secuer', 'mycofix secure'] },
  { name: 'ام- موبيلايز - M-Mobilize', driveCode: '431', jdeCode: '1018632', aliases: ['ام موبيلايز', 'm-mobilize'] },
  { name: 'كيومتيوكس', driveCode: '671', jdeCode: '1259049' },
  { name: 'مضاد سموم ايجافيت يونيك بلس', driveCode: '688', jdeCode: '1281422' },
  { name: 'سوبر اميونى Super imuni', driveCode: '634', jdeCode: '1178960' },
  { name: 'سوبر اميونى - Super imuni (petology)', driveCode: '683', jdeCode: '683' },
  { name: 'ايه دى اسد بيوتى بلص Ad Acid Butyplus', driveCode: '629', jdeCode: '1512673' },
  { name: 'روفى يسيت', driveCode: '686', jdeCode: '1316357' },
  { name: 'مضاد سموم ميجا بايو', driveCode: '687', jdeCode: '1324963' },
  { name: 'فيلاكس بودرة Felix powder', driveCode: '633', jdeCode: '1178951' },
  { name: 'فيلاكس بودرة - Felix powder (petology)', driveCode: '682', jdeCode: '682' },
  { name: 'املاح تسمين 2 Premix Mineral Broiler DSM', driveCode: '613', jdeCode: '' },
  { name: 'لايسين - Lysine HCL', shortName: 'لايسين', driveCode: '304', jdeCode: '13556', aliases: ['لايسين', 'lysine', 'lysine hcl'] },
  { name: 'ثرونين - L-Threonine', shortName: 'L-Threonine', driveCode: '302', jdeCode: '13557', aliases: ['ثرونين', 'l-threonine', 'threonine'] },
  { name: 'ميثونين - L Methionin', shortName: 'ميثونين', driveCode: '303', jdeCode: '13558', aliases: ['ميثونين', 'ميثيونين', 'l methionin', 'methionine'] },
  { name: 'تربتوفان L Tryptophan', driveCode: '505', jdeCode: '16348' },
  { name: 'ناتشورا ابيتايين ال', driveCode: '569', jdeCode: '907921' },
  { name: 'تيورين - taurine', driveCode: '650', jdeCode: '1209893' },
  { name: 'تاورين - taurine (petology)', driveCode: '681', jdeCode: '681' },
  { name: 'ناتشورا ابيتايين ال', driveCode: '690', jdeCode: '1283006' },
  { name: 'ميثونين سائل - L Methionin', driveCode: '693', jdeCode: '1287795', aliases: ['liquid methonin', 'liquid methionine', 'ميثونين سائل'] },
  { name: 'macro madura vit', driveCode: '694', jdeCode: '1303943' },
  { name: 'دايكلا زوريل macro diclo vit', driveCode: '696', jdeCode: '1303935' },
  { name: 'ديكلا شمس - Dicla Shams', driveCode: '322', jdeCode: '5268' },
  { name: 'مضاد اكسده OX Anti', driveCode: '324', jdeCode: '1472106', aliases: ['anti ox', 'ox anti', 'مضاد اكسده'] },
  { name: 'مادورا فارما - Madura pharm', driveCode: '333', jdeCode: '939704' },
  { name: 'مادورا شمس - MaduraShams', driveCode: '332', jdeCode: '970331' },
  { name: 'مخلوط إضافات موحد Premix 8-10 A', driveCode: '848', jdeCode: '1088745' },
  { name: 'كلوبى COLIBDOL', driveCode: '593', jdeCode: '995749' },
  { name: 'سالينو شمس Salino shams', driveCode: '595', jdeCode: '894294' },
  { name: 'اوبى كوكس Opi-cox', driveCode: '321', jdeCode: '907913' },
  { name: 'ماكسيبان', driveCode: '720', jdeCode: '16346' },
  { name: 'ماكسيبان النمر Customer addition (Maxibin)', driveCode: '614', jdeCode: '614' },
  { name: 'ماكسيبان الهنا Customer addition (Maxibin)', driveCode: '625', jdeCode: '625' },
  { name: 'ماكسيبان تيلاج Customer addition (Maxibin)', driveCode: '626', jdeCode: '626' },
  { name: 'بروفورس اس ار 130', driveCode: '656', jdeCode: '656' },
  { name: 'Customer addition (Monteban)', driveCode: '617', jdeCode: '617' },
  { name: 'Customer addition (الانكوبان النمر)', driveCode: '620', jdeCode: '620' },
  { name: 'القنطرة customer addition HY-D', driveCode: '695', jdeCode: '695' },
  { name: 'مادورا شمس 1% 25 كجم MaduraShams', driveCode: '721', jdeCode: '1333456' },
  { name: 'بيب بودرة P.E.P', driveCode: '311', jdeCode: '13797' },
  { name: 'اورجستيم بودرة مستورد', driveCode: '638', jdeCode: '13798', aliases: ['اورجستيم بودرة مستورد', 'origo stim', 'origostim'] },
  { name: 'ايفاباكت جرانيول لينكومايسين 11%', driveCode: '639', jdeCode: '1295998' },
  { name: 'مادورا شمس 1% 1 ك MaduraShams', driveCode: '640', jdeCode: '918461' },
  { name: 'لينكومايسين 11%', driveCode: '644', jdeCode: '1338417' },
  { name: 'لينكومايسين 4,4', driveCode: '645', jdeCode: '1338409' },
  { name: 'دايكوميد - دايكلا زوريل (Dicomed)', driveCode: '646', jdeCode: '1345191' },
  { name: 'مضاد اكسده سائل TBHQ', driveCode: '642', jdeCode: '1196121' },
  { name: 'إندوكس - Anti oxidint', driveCode: '319', jdeCode: '16739' },
  { name: 'مضاد اكسدة Mintal Antioxy', driveCode: '643', jdeCode: '1314327' },
  { name: 'مضاد اكسده Antiox Guard pro', driveCode: '648', jdeCode: '1348084' },
  { name: 'ديكلا شمس 1 - Dicla Shams', driveCode: '323', jdeCode: '751437', aliases: ['ديكلا شمس 1', 'dicla shams 1', 'dicla shams'] },
  { name: 'رونزيم هاي فوس - Hi-Phos سائل 20000', driveCode: '677', jdeCode: '1155178' },
  { name: 'micro tech 20000 (فايتيز)', driveCode: '661', jdeCode: '1241148' },
  { name: 'micro tech 10000 (فايتيز)', driveCode: '375', jdeCode: '1004880' },
  { name: 'رونزيم هاي فوس - Hi-Phos', driveCode: '325', jdeCode: '13563' },
  { name: 'رونزايم برواكت - Ronozyme Proact', driveCode: '318', jdeCode: '13565' },
  { name: 'انزيم بروتيز 100000', driveCode: '653', jdeCode: '1093413' },
  { name: 'انزيم بروتيز 20000', driveCode: '675', jdeCode: '1248876' },
  { name: 'رونزيم هاي فوس - Hi-Phos 20000 بودر', driveCode: '678', jdeCode: '1287242', aliases: ['hi-phos gt 20000', 'hi-phos gt', 'hi phos gt', 'hi-phos 20000'] },
  { name: 'وينوبرو بروتيز 60000', driveCode: '679', jdeCode: '1319283' },
  { name: 'ماديوارسين', driveCode: '692', jdeCode: '1375162' },
  { name: 'انسيبرا بلاس - Enspira', driveCode: '485', jdeCode: '1014914' },
  { name: 'انزيم كومباوند - Compound Enzyme', driveCode: '452', jdeCode: '1024881' },
  { name: 'كورن', driveCode: '589', jdeCode: '1031714' },
  { name: 'انزيم بيتا جلوكانيز', driveCode: '603', jdeCode: '1104590' },
  { name: 'انزيم بيتا ماننيز - Mananase Enzyme 20000', driveCode: '621', jdeCode: '1110076' },
  { name: 'انزيم ليبيز 50000 - (50)Lipase', driveCode: '622', jdeCode: '1024830' },
  { name: 'انزيم بيتا ماننيز - Mananase Enzyme 50000', driveCode: '631', jdeCode: '1110084' },
  { name: 'رونزيم - Ronozyme WX 2000', driveCode: '317', jdeCode: '862575' },
  { name: 'رونزيم - Ronozyme VP', driveCode: '337', jdeCode: '13857' },
  { name: 'رونزيم - Ronozyme WX', driveCode: '676', jdeCode: '13858', aliases: ['ronzyme wx', 'ronzyme wx 2000', 'ronozyme wx 2000'] },
  { name: 'أنزيم جلكوز إكسيديز 20000', driveCode: '632', jdeCode: '1246141' },
  { name: 'أنزيم جلكوز إكسيديز 10000', driveCode: '636', jdeCode: '1339153' },
  { name: 'mastercube plus', driveCode: '701', jdeCode: '1461941' },
  { name: 'مادورا شمس 1% 10 كجم MaduraShams', driveCode: '710', jdeCode: '907657' },
  { name: 'ملح طعام salt', shortName: 'ملح طعام', driveCode: '300', jdeCode: '13569', aliases: ['ملح طعام', 'salt'] },
  { name: 'بيكربونات صوديوم - Sodium Bicarbonate', shortName: 'بيكربونات صوديوم', driveCode: '301', jdeCode: '13572', aliases: ['بيكربونات صوديوم', 'sodium bicarbonate'] },
  { name: 'سلفات صوديوم (petology)', driveCode: '702', jdeCode: '702' },
  { name: 'مادورا شمس 10 كجم MaduraShams', driveCode: '711', jdeCode: '1393838' },
  { name: 'yucca powder', driveCode: '750', jdeCode: '750', aliases: ['yucca powder', 'yucca', 'يوكا'] },
  { name: 'رونيون (بايل اسيد) - Bile Acid Runnen', shortName: 'رونيون (بايل اسيد ) - Bile Acid Runnen', driveCode: '331', jdeCode: '907948', aliases: ['رونيون (بايل اسيد )', 'bile acid runnen', 'bile acid 200g', 'bile acid'] },
  { name: 'فولاميل اكسترا بودرة Volamil Emulsifier', driveCode: '598', jdeCode: '868571' },
  { name: 'Intest- Plus Quattro 70 %', driveCode: '672', jdeCode: '1263187' },
  { name: 'Intest- Plus Quattro 70 % (Petology)', driveCode: '680', jdeCode: '680' },
  { name: 'بريمكس كلاب puppy', driveCode: '845', jdeCode: '1178898' },
  { name: 'Adult Dog Premix', driveCode: '846', jdeCode: '1178898' },
  { name: 'Adult Cat Premix', driveCode: '847', jdeCode: '1220136' },
  { name: 'كيروكس اتش تى Careox Ht', driveCode: '601', jdeCode: '1049439' },
  { name: 'تراى بيوترين BTG', driveCode: '602', jdeCode: '1104581' },
  { name: 'تيلو شمس', driveCode: '604', jdeCode: '1347508' },
  { name: 'كلورا شمس 20%', driveCode: '605', jdeCode: '17160' },
  { name: 'مخلوط إضافات موحد Premix 26-5 A', driveCode: '844', jdeCode: '1088745' },
  { name: 'مخلوط إضافات موحد Premix 15-4 B', driveCode: '843', jdeCode: '1088745' },
  { name: 'امنيل بروبيوتك Amnil', driveCode: '391', jdeCode: '1010622' },
  { name: 'نوفيلا 100 Novella 100', driveCode: '421', jdeCode: '1016071' },
  { name: 'بيوترونيك توب سائل 10 كجم', driveCode: '450', jdeCode: '13781' },
  { name: 'مخلوط اضافات موحد Premix 15-4', driveCode: '842', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد Premix 26-2', driveCode: '840', jdeCode: '1088745' },
  { name: 'بيوتين dsm (ROVIMIX BIOTIN - 20KG)', driveCode: '422', jdeCode: '13616' },
  { name: 'فيتامين A1000 (ROVIMIX A1000 - 20KG)', driveCode: '423', jdeCode: '13610' },
  { name: 'فيتامين B2 (ROVIMIX B2 80 SD - 20KG)', driveCode: '424', jdeCode: '13613' },
  { name: 'فيتامينات نياسين مستورد (ROVIMIX NIACIN - 25KG)', driveCode: '425', jdeCode: '13608' },
  { name: 'سالينومايسين', driveCode: '426', jdeCode: '1485249' },
  { name: 'فيتامين B1 م (ROVIMIX B1 - 25KG)', driveCode: '427', jdeCode: '13611' },
  { name: 'فيتامين B12 1% (VITAMIN B12 1% FEED GRADE - 25KG)', driveCode: '428', jdeCode: '13612' },
  { name: 'بيترونيك توب 3 - Biotrinic Top3', shortName: 'Biotrinic TOP 3', driveCode: '329', jdeCode: '13855', aliases: ['biotrinic top 3', 'biotrinic top3', 'بيوترونيك توب 3'] },
  { name: 'روفيميكس بياض', driveCode: '839', jdeCode: '' },
  { name: 'مخلوط موحد كلاب Adult Dog Premix', driveCode: '838', jdeCode: '1219071' },
  { name: 'ياموس بلس', driveCode: '900', jdeCode: '', aliases: ['ياموس بلس', 'yeamos plus', 'yamos plus'] },
  { name: 'مخلوط اضافات موحد Premix 31-10', driveCode: '837', jdeCode: '1088745' },
  { name: 'بريمكس قطط Cat Premix (B2) IR', driveCode: '855', jdeCode: '1220136' },
  { name: 'كامينا كارموازين احمر بودر', driveCode: '647', jdeCode: '1216989' },
  { name: 'مخلوط موحد قطط Adult Cat Premix IR', driveCode: '652', jdeCode: '1220566' },
  { name: 'بريمكس هاي دي (Rovimix HY-D)', driveCode: '836', jdeCode: '1304090', aliases: ['hy-d', 'rovimix hy-d', 'بريمكس هاي دي'] },
  { name: 'بريمكس بط تسمين - Duck Premix 2 KG', driveCode: '358', jdeCode: '868352' },
  { name: 'بريمكس بط امهات - Duck Premix 2 KG', driveCode: '363', jdeCode: '868361' },
  { name: 'بريمكس سمك - Fish Permix', driveCode: '415', jdeCode: '868379', aliases: ['اضافات سمك', 'fish premix', 'fish permix', 'بريمكس سمك'] },
  { name: 'بريمكس تسمين ابيض بادي نامى تسويقى - Starter Grower Premix', driveCode: '588', jdeCode: '807154' },
  { name: 'بريمكس كلاب محلى Dog Premix (B)', driveCode: '635', jdeCode: '1178898' },
  { name: 'مخلوط موحد كلاب (Adult 20-5-E)', driveCode: '651', jdeCode: '1219071' },
  { name: 'مخلوط موحد قطط Adult Cat 25-10', driveCode: '654', jdeCode: '1220566' },
  { name: 'كاروفيل 30/10 صفار بيض 12', driveCode: '655', jdeCode: '' },
  { name: 'مخلوط موحد قطط Adult Cat 25-10A/1', driveCode: '684', jdeCode: '1220566' },
  { name: 'مخلوط موحد كلاب (Adult Dog 19-12A)', driveCode: '685', jdeCode: '1219071' },
  { name: 'مخلوط اضافات موحد Premix 27/12', driveCode: '689', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد Premix 8/10/1', driveCode: '691', jdeCode: '1088745' },
  { name: 'مخلوط إضافات موحد Premix 27-10 A', driveCode: '800', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد 27/10A[linco] Premix', driveCode: '801', jdeCode: '1488706' },
  { name: 'مخلوط اضافات موحد Premix 27/12/4', driveCode: '802', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد Premix 27/12/5', driveCode: '803', jdeCode: '1088745' },
  { name: 'مخلوط اضافات اعلاف لمرحلة البادى ST 7-1 A', driveCode: '804', jdeCode: '1499780' },
  { name: 'مخلوط اضافات اعلاف لمرحلة النامى GR 7-1 A', driveCode: '805', jdeCode: '1499798' },
  { name: 'مخلوط اضافات موحد A 3-7 Premix', driveCode: '806', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد Premix 27/12/9', driveCode: '807', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد Premix 27/12/10', driveCode: '808', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد Premix 27/12/11', driveCode: '809', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد 19-9 A', driveCode: '810', jdeCode: '1088745', aliases: ['premix 19-9 a', '19-9 a'] },
  { name: 'مخلوط اضافات موحد Premix 13-2', driveCode: '811', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد Premix 13-2/1', driveCode: '812', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد Premix 13-2/2', driveCode: '813', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد Premix 5-3', driveCode: '814', jdeCode: '1088745' },
  { name: 'بريمكس قطط Cat Premix (B)', driveCode: '815', jdeCode: '1220136' },
  { name: 'مخلوط اضافات موحد Premix 5-3.1', driveCode: '816', jdeCode: '1088745' },
  { name: 'مخلوط اضافات اعلاف لمرحلة النامى GR 17-2 A', driveCode: '817', jdeCode: '1499798' },
  { name: 'مخلوط إضافات موحد Premix 20-9 A', driveCode: '823', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد Premix 5-3.4', driveCode: '819', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد prlmex 8_10/1', driveCode: '820', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد Premix 5-3.6', driveCode: '821', jdeCode: '1088745' },
  { name: 'مخلوط موحد كلاب puppy28-4', driveCode: '822', jdeCode: '1219071' },
  { name: 'مخلوط اضافات اعلاف لمرحلة الناهى linco 17-2', driveCode: '850', jdeCode: '1488706' },
  { name: 'مخلوط اضافات اعلاف 17-2 /1', driveCode: '824', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد Premix 5-3.9', driveCode: '825', jdeCode: '1088745' },
  { name: 'رونزيم برواكت PREMIX 1KG', driveCode: '852', jdeCode: '1503970' },
  { name: 'rovimix tilapia lite', driveCode: '827', jdeCode: '1508658' },
  { name: 'مخلوط اضافات موحد A1 4.13', driveCode: '828', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد 13-4 A', driveCode: '829', jdeCode: '1088745' },
  { name: 'بريمكس قطط Kitten Premix', driveCode: '830', jdeCode: '1220136' },
  { name: 'tilapia px .02%', driveCode: '831', jdeCode: '1518047' },
  { name: 'مخلوط اضافات اعلاف لمرحلة الناهى 17-2', driveCode: '851', jdeCode: '1088745' },
  { name: 'مخلوط اضافات اعلاف لمرحلة الناهى FN 17-2 A', driveCode: '833', jdeCode: '1488706' },
  { name: 'مخلوط اضافات موحد Premix 1-9', driveCode: '834', jdeCode: '1088745', aliases: ['مخلوط اضافات موحد 9-1 premix', 'مخلوط اضافات موحد 1-9 premix', 'premix 1-9', 'premix 9-1'] },
  { name: 'مخلوط موحد كلاب (Adult Dog 10-8A)', driveCode: '835', jdeCode: '1219071' },
  { name: 'رونزيم هاي فوس + مالتى جرين', driveCode: '666', jdeCode: '1514599' },
  { name: 'مخلوط اضافات موحد A1 3-7 Premix', driveCode: '890', jdeCode: '1088745' },
  { name: 'مخلوط اضافات موحد Premix 7-3 A/2', driveCode: '891', jdeCode: '1088745' },
  { name: 'Ronozyme Proact (CT) 1000 kg', driveCode: '892', jdeCode: '1513596' },
  { name: 'مخلوط اضافات موحد Premix7-3 B', driveCode: '893', jdeCode: '1088745' },
  { name: 'كود جديد adult cat premix', driveCode: '2001', jdeCode: '1519701' },
  { name: 'كود جديد kitten premix', driveCode: '2002', jdeCode: '1519710' },
  { name: 'كود جديد puppy premix', driveCode: '2003', jdeCode: '1519728' },
  { name: 'كود جديد adult dog premix', driveCode: '2004', jdeCode: '1519736', aliases: ['adult dog premix'] },
  { name: 'Fermented assistant plus', driveCode: '2005', jdeCode: '1535429' },
  { name: 'Lecithin liquid crude', driveCode: '2006', jdeCode: '1535437' },
  { name: 'Cholesterol 99%', driveCode: '2007', jdeCode: '1535445' },
  { name: 'مخلوط اضافات اعلاف لمرحلة الناهى FN 27-10 A', driveCode: '849', jdeCode: '1488706' },
];

export const DRIVE_OVERRIDES_STORAGE_KEY = 'control_drive_code_overrides_v1';
export const DRIVE_CUSTOM_CATALOG_STORAGE_KEY = 'control_drive_custom_catalog_v1';

/**
 * Normalizes Arabic & English text for reliable matching across variations.
 */
export function normalizeItemText(text: string): string {
  return (text || '')
    .toString()
    .toLowerCase()
    // Remove Arabic diacritics
    .replace(/[\u064B-\u065F\u0670]/g, '')
    // Normalize Arabic letters
    .replace(/[أإآ]/g, 'ا')
    .replace(/ة/g, 'ه')
    .replace(/ى/g, 'ي')
    .replace(/ؤ/g, 'و')
    .replace(/ئ/g, 'ي')
    // Common English/Arabic spelling variants
    .replace(/\bpremix\b|\bpermix\b|\bpermex\b|\bprlmex\b/g, 'premix')
    .replace(/\bmethionin\b|\bmethionine\b|\bmethonin\b/g, 'methionin')
    .replace(/\bsecuer\b|\bsecure\b/g, 'secuer')
    .replace(/\byeamos\b|\byamos\b/g, 'yeamos')
    .replace(/\bbiotrinic\b|\bbitronic\b/g, 'biotrinic')
    .replace(/\bronzyme\b/g, 'ronozyme')
    .replace(/\bhi[\s-]*phos\b/g, 'hi phos')
    // Replace punctuation with spaces
    .replace(/[-_/\\()[\]{}.,:+*]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Items that must be grouped BY NAME even when the report lists them as variants
 * (e.g. "Corn" + "Corn B", "Local soya 46%" + "SOYA 46% B").
 * Add more base names here (normalized form, lower-case) to group other items the same way.
 */
export const NAME_GROUPED_BASE_ITEMS: string[] = [
  'corn',
  'yellow corn',
  'soya 46%',
  'soya 44%',
];

/**
 * Strips variant markers ("Local" prefix before soya, trailing standalone "B") ONLY when the
 * result is one of NAME_GROUPED_BASE_ITEMS, so unrelated items such as "Premix 15-4 B" stay separate.
 * Returns the display name to use and whether the item belongs to a name-grouped family.
 */
export function groupItemByName(rawName: string): { name: string; grouped: boolean } {
  const original = (rawName || '').toString().trim();
  if (!original) return { name: original, grouped: false };

  let stripped = original
    .replace(/^\s*(?:local|محلي|محلى)\s+(?=soya|soy|صويا)/i, '')
    .replace(/[\s\-(]+b\s*\)?\s*$/i, '')
    .replace(/\s+/g, ' ')
    .trim();

  const bases = NAME_GROUPED_BASE_ITEMS.map((b) => normalizeItemText(b));
  if (bases.includes(normalizeItemText(stripped))) {
    return { name: stripped, grouped: true };
  }
  // Already a base name without any variant marker (e.g. plain "Corn")
  if (bases.includes(normalizeItemText(original))) {
    return { name: original, grouped: true };
  }
  return { name: original, grouped: false };
}

function tokenize(text: string): string[] {
  const norm = normalizeItemText(text);
  if (!norm) return [];
  return norm.split(' ').filter((t) => t.length > 0);
}

function tokenOverlapScore(queryName: string, candidateStrings: string[]): number {
  const qTokens = tokenize(queryName);
  if (qTokens.length === 0) return 0;

  let bestScore = 0;
  for (const cand of candidateStrings) {
    const normCand = normalizeItemText(cand);
    if (!normCand) continue;
    if (normCand === normalizeItemText(queryName)) return 1.0;

    const cTokens = tokenize(cand);
    if (cTokens.length === 0) continue;

    let matches = 0;
    for (const qt of qTokens) {
      if (cTokens.includes(qt) || cTokens.some((ct) => ct.length >= 3 && qt.length >= 3 && (ct.includes(qt) || qt.includes(ct)))) {
        matches++;
      }
    }

    // Different formula/percentage numbers (e.g. "1-9" vs "15-4") must not cross-match
    const qNums = qTokens.filter((t) => /^\d+$/.test(t));
    const cNums = cTokens.filter((t) => /^\d+$/.test(t));
    let numBonus = 0;
    if (qNums.length > 0 && cNums.length > 0) {
      const numMatches = qNums.filter((n) => cNums.includes(n)).length;
      if (numMatches === qNums.length) {
        numBonus = 0.25;
      } else if (numMatches === 0) {
        continue;
      }
    }

    const score = matches / Math.max(qTokens.length, 1) + numBonus;
    if (score > bestScore) {
      bestScore = score;
    }
  }
  return bestScore;
}

function deriveShortDriveName(entry: DriveCodeEntry): string {
  if (entry.shortName) return entry.shortName;
  const raw = (entry.name || '').trim();
  const dashIdx = raw.indexOf(' - ');
  if (dashIdx !== -1) {
    const left = raw.slice(0, dashIdx).trim();
    if (/[\u0600-\u06FF]/.test(left)) {
      return left;
    }
  }
  return raw;
}

export interface DriveCatalogMatchInfo {
  driveCode: string;
  shortName: string;
  fullName: string;
  catalogOrder: number;
}

export function lookupDriveCatalogInfo(
  itemName: string,
  jdeCode: string,
  overrides: Record<string, string> = {},
  customCatalog: DriveCodeEntry[] = []
): DriveCatalogMatchInfo {
  const cleanName = (itemName || '').trim();
  const cleanJde = (jdeCode || '').trim();
  const normName = normalizeItemText(cleanName);

  const combinedCatalog = [...customCatalog, ...MASTER_DRIVE_CODES_CATALOG];

  const findEntryByDriveCode = (codeStr: string): { entry: DriveCodeEntry; index: number } | null => {
    const idx = combinedCatalog.findIndex((e) => e.driveCode === codeStr.trim());
    return idx !== -1 ? { entry: combinedCatalog[idx], index: idx } : null;
  };

  // 1. Check manual user overrides
  const overrideKey1 = `${cleanJde}::${normName}`;
  let overrideDriveCode: string | undefined;
  if (overrides[overrideKey1] !== undefined) overrideDriveCode = overrides[overrideKey1];
  else if (cleanName && overrides[cleanName] !== undefined) overrideDriveCode = overrides[cleanName];
  else if (normName && overrides[`NAME::${normName}`] !== undefined) overrideDriveCode = overrides[`NAME::${normName}`];

  if (overrideDriveCode !== undefined) {
    const hit = overrideDriveCode ? findEntryByDriveCode(overrideDriveCode) : null;
    return {
      driveCode: overrideDriveCode,
      shortName: hit ? deriveShortDriveName(hit.entry) : cleanName,
      fullName: hit ? hit.entry.name : cleanName,
      catalogOrder: hit ? hit.index : Number.MAX_SAFE_INTEGER,
    };
  }

  // 2a. Exact normalized name or alias match (prefer matching JDE too if both match)
  const exactMatches: { entry: DriveCodeEntry; index: number }[] = [];
  combinedCatalog.forEach((entry, index) => {
    const namesToCheck = [entry.name, ...(entry.aliases || [])];
    if (namesToCheck.some((n) => normalizeItemText(n) === normName)) {
      exactMatches.push({ entry, index });
    }
  });
  if (exactMatches.length > 0) {
    let chosen = exactMatches[0];
    if (cleanJde) {
      const jdeHit = exactMatches.find((e) => e.entry.jdeCode === cleanJde);
      if (jdeHit) chosen = jdeHit;
    }
    return {
      driveCode: chosen.entry.driveCode,
      shortName: deriveShortDriveName(chosen.entry),
      fullName: chosen.entry.name,
      catalogOrder: chosen.index,
    };
  }

  const hasVariantMismatch = (query: string, entry: DriveCodeEntry): boolean => {
    const qTokens = tokenize(query);
    if (qTokens.length === 0) return false;
    const entryTokens = new Set<string>();
    for (const str of [entry.name, ...(entry.aliases || [])]) {
      for (const tok of tokenize(str)) entryTokens.add(tok);
    }
    const distinguishingTokens = ['b', 'dark', 'ukrainian', 'imported', 'white', 'select', 'secure', 'lite'];
    return distinguishingTokens.some((dt) => qTokens.includes(dt) && !entryTokens.has(dt));
  };

  // 2b. Filter by JDE code if provided (reject candidates with conflicting variant tokens like "Corn B" vs "Corn")
  if (cleanJde) {
    const jdeCandidates = combinedCatalog
      .map((entry, index) => ({ entry, index }))
      .filter((item) => item.entry.jdeCode === cleanJde && !hasVariantMismatch(cleanName, item.entry));
    if (jdeCandidates.length === 1) {
      const chosen = jdeCandidates[0];
      return {
        driveCode: chosen.entry.driveCode,
        shortName: deriveShortDriveName(chosen.entry),
        fullName: chosen.entry.name,
        catalogOrder: chosen.index,
      };
    }
    if (jdeCandidates.length > 1) {
      // Disambuate shared JDE codes (e.g. 11342, 1088745, 752042, 1219071) by name similarity
      let chosen: { entry: DriveCodeEntry; index: number } | null = null;
      let bestScore = 0;
      for (const cand of jdeCandidates) {
        const score = tokenOverlapScore(cleanName, [cand.entry.name, ...(cand.entry.aliases || [])]);
        if (score > bestScore) {
          bestScore = score;
          chosen = cand;
        }
      }
      if (chosen && bestScore >= 0.45) {
        return {
          driveCode: chosen.entry.driveCode,
          shortName: deriveShortDriveName(chosen.entry),
          fullName: chosen.entry.name,
          catalogOrder: chosen.index,
        };
      }
    }
  }

  // 2c. Name-based matching when JDE code is missing or not in catalog
  let bestMatch: { entry: DriveCodeEntry; index: number } | null = null;
  let highestScore = 0;

  combinedCatalog.forEach((entry, index) => {
    if (hasVariantMismatch(cleanName, entry)) return;
    const namesToCheck = [entry.name, ...(entry.aliases || [])];
    for (const candidate of namesToCheck) {
      const normCand = normalizeItemText(candidate);
      if (!normCand || !normName) continue;
      if (normName.length >= 4 && (normCand === normName || normCand.includes(normName) || normName.includes(normCand))) {
        const lenRatio = Math.min(normName.length, normCand.length) / Math.max(normName.length, normCand.length);
        const score = 0.75 + 0.25 * lenRatio;
        if (score > highestScore) {
          highestScore = score;
          bestMatch = { entry, index };
        }
      }
    }

    const tokenScore = tokenOverlapScore(cleanName, namesToCheck);
    if (tokenScore > highestScore) {
      highestScore = tokenScore;
      bestMatch = { entry, index };
    }
  });

  if (bestMatch && highestScore >= 0.65) {
    const chosen: { entry: DriveCodeEntry; index: number } = bestMatch;
    return {
      driveCode: chosen.entry.driveCode,
      shortName: deriveShortDriveName(chosen.entry),
      fullName: chosen.entry.name,
      catalogOrder: chosen.index,
    };
  }

  return {
    driveCode: '',
    shortName: cleanName,
    fullName: cleanName,
    catalogOrder: Number.MAX_SAFE_INTEGER,
  };
}

/**
 * Looks up "كود الدريف" using:
 * 1. User manual overrides (keyed by code::name or normalized name)
 * 2. Custom imported catalog + Built-in master catalog
 */
export function lookupDriveCode(
  itemName: string,
  jdeCode: string,
  overrides: Record<string, string> = {},
  customCatalog: DriveCodeEntry[] = []
): string {
  return lookupDriveCatalogInfo(itemName, jdeCode, overrides, customCatalog).driveCode;
}

/**
 * Parses an uploaded Excel (.xlsx/.xls) or CSV file containing item codes
 * into DriveCodeEntry[] so the user can extend or refresh the catalog anytime.
 */
export async function parseDriveCodesFile(file: File): Promise<DriveCodeEntry[]> {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array' });
  const firstSheetName = wb.SheetNames[0];
  if (!firstSheetName) return [];

  const ws = wb.Sheets[firstSheetName];
  const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (rows.length < 2) return [];

  // Detect header indices
  let headerRowIdx = 0;
  let nameIdx = -1;
  let driveIdx = -1;
  let jdeIdx = -1;

  for (let r = 0; r < Math.min(rows.length, 5); r++) {
    const row = rows[r].map((c) => String(c || '').trim().toLowerCase());
    const nI = row.findIndex((h) => h.includes('اسم') || h.includes('إسم') || h.includes('name'));
    const dI = row.findIndex((h) => h.includes('دريف') || h.includes('drive'));
    const jI = row.findIndex((h) => h.includes('jde') || h.includes('كود الصنف'));
    if (nI !== -1 && dI !== -1) {
      headerRowIdx = r;
      nameIdx = nI;
      driveIdx = dI;
      jdeIdx = jI !== -1 && jI !== dI ? jI : -1;
      break;
    }
  }

  // Fallback column order matching screenshot: [اسم الصنف, كود الصنف دريف, كود الصنف JDE]
  if (nameIdx === -1) nameIdx = 0;
  if (driveIdx === -1) driveIdx = 1;
  if (jdeIdx === -1) jdeIdx = 2;

  const entries: DriveCodeEntry[] = [];
  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!Array.isArray(row)) continue;
    const name = String(row[nameIdx] ?? '').trim();
    const driveCode = String(row[driveIdx] ?? '').trim();
    const jdeCode = String(row[jdeIdx] ?? '').trim();
    if (name && driveCode && driveCode !== '-') {
      entries.push({ name, driveCode, jdeCode });
    }
  }

  return entries;
}
