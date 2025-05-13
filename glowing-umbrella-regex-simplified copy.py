import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import subprocess
import re
# Removed: import docx # No longer needed

VOICES = [
    # English (en_GB)
    ("en_GB: alan (low)", "piper/piper-voices/en/en_GB/alan/low/en_GB-alan-low.onnx"),
    ("en_GB: alan (medium)", "piper/piper-voices/en/en_GB/alan/medium/en_GB-alan-medium.onnx"),
    ("en_GB: alba (medium)", "piper/piper-voices/en/en_GB/alba/medium/en_GB-alba-medium.onnx"),
    ("en_GB: aru (medium)", "piper/piper-voices/en/en_GB/aru/medium/en_GB-aru-medium.onnx"),
    ("en_GB: cori (medium)", "piper/piper-voices/en/en_GB/cori/medium/en_GB-cori-medium.onnx"),
    ("en_GB: cori (high)", "piper/piper-voices/en/en_GB/cori/high/en_GB-cori-high.onnx"),
    ("en_GB: jenny_dioco (medium)", "piper/piper-voices/en/en_GB/jenny_dioco/medium/en_GB-jenny_dioco-medium.onnx"),
    ("en_GB: northern_english_male (medium)", "piper/piper-voices/en/en_GB/northern_english_male/medium/en_GB-northern_english_male-medium.onnx"),
    ("en_GB: semaine (medium)", "piper/piper-voices/en/en_GB/semaine/medium/en_GB-semaine-medium.onnx"),
    ("en_GB: southern_english_female (low)", "piper/piper-voices/en/en_GB/southern_english_female/low/en_GB-southern_english_female-low.onnx"),
    ("en_GB: vctk (medium)", "piper/piper-voices/en/en_GB/vctk/medium/en_GB-vctk-medium.onnx"),

    # English (en_US)
    ("en_US: amy (low)", "piper/piper-voices/en/en_US/amy/low/en_US-amy-low.onnx"),
    ("en_US: amy (medium)", "piper/piper-voices/en/en_US/amy/medium/en_US-amy-medium.onnx"),
    ("en_US: arctic (medium)", "piper/piper-voices/en/en_US/arctic/medium/en_US-arctic-medium.onnx"),
    ("en_US: bryce (medium)", "piper/piper-voices/en/en_US/bryce/medium/en_US-bryce-medium.onnx"),
    ("en_US: danny (low)", "piper/piper-voices/en/en_US/danny/low/en_US-danny-low.onnx"),
    ("en_US: hfc_female (medium)", "piper/piper-voices/en/en_US/hfc_female/medium/en_US-hfc_female-medium.onnx"),
    ("en_US: hfc_male (medium)", "piper/piper-voices/en/en_US/hfc_male/medium/en_US-hfc_male-medium.onnx"),
    ("en_US: joe (medium)", "piper/piper-voices/en/en_US/joe/medium/en_US-joe-medium.onnx"),
    ("en_US: john (medium)", "piper/piper-voices/en/en_US/john/medium/en_US-john-medium.onnx"),
    ("en_US: kathleen (low)", "piper/piper-voices/en/en_US/kathleen/low/en_US-kathleen-low.onnx"),
    ("en_US: kristin (medium)", "piper/piper-voices/en/en_US/kristin/medium/en_US-kristin-medium.onnx"),
    ("en_US: kusal (medium)", "piper/piper-voices/en/en_US/kusal/medium/en_US-kusal-medium.onnx"),
    ("en_US: l2arctic (medium)", "piper/piper-voices/en/en_US/l2arctic/medium/en_US-l2arctic-medium.onnx"),
    ("en_US: lessac (low)", "piper/piper-voices/en/en_US/lessac/low/en_US-lessac-low.onnx"),
    ("en_US: lessac (medium)", "piper/piper-voices/en/en_US/lessac/medium/en_US-lessac-medium.onnx"),
    ("en_US: lessac (high)", "piper/piper-voices/en/en_US/lessac/high/en_US-lessac-high.onnx"),
    ("en_US: libritts (high)", "piper/piper-voices/en/en_US/libritts/high/en_US-libritts-high.onnx"),
    ("en_US: libritts_r (medium)", "piper/piper-voices/en/en_US/libritts_r/medium/en_US-libritts_r-medium.onnx"),
    ("en_US: ljspeech (medium)", "piper/piper-voices/en/en_US/ljspeech/medium/en_US-ljspeech-medium.onnx"),
    ("en_US: ljspeech (high)", "piper/piper-voices/en/en_US/ljspeech/high/en_US-ljspeech-high.onnx"),
    ("en_US: norman (medium)", "piper/piper-voices/en/en_US/norman/medium/en_US-norman-medium.onnx"),
    ("en_US: reza_ibrahim (medium)", "piper/piper-voices/en/en_US/reza_ibrahim/medium/en_US-reza_ibrahim-medium.onnx"),
    ("en_US: ryan (low)", "piper/piper-voices/en/en_US/ryan/low/en_US-ryan-low.onnx"),
    ("en_US: ryan (medium)", "piper/piper-voices/en/en_US/ryan/medium/en_US-ryan-medium.onnx"),
    ("en_US: ryan (high)", "piper/piper-voices/en/en_US/ryan/high/en_US-ryan-high.onnx"),
    ("en_US: sam (medium)", "piper/piper-voices/en/en_US/sam/medium/en_US-sam-medium.onnx"),

    # Spanish (es_ES/es_MX)
    ("es_ES: carlfm (x_low)", "piper/piper-voices/es/es_ES/carlfm/x_low/es_ES-carlfm-x_low.onnx"),
    ("es_ES: davefx (medium)", "piper/piper-voices/es/es_ES/davefx/medium/es_ES-davefx-medium.onnx"),
    ("es_ES: mls_10246 (low)", "piper/piper-voices/es/es_ES/mls_10246/low/es_ES-mls_10246-low.onnx"),
    ("es_ES: mls_9972 (low)", "piper/piper-voices/es/es_ES/mls_9972/low/es_ES-mls_9972-low.onnx"),
    ("es_ES: sharvard (medium)", "piper/piper-voices/es/es_ES/sharvard/medium/es_ES-sharvard-medium.onnx"),
    ("es_MX: ald (medium)", "piper/piper-voices/es/es_MX/ald/medium/es_MX-ald-medium.onnx"),
    ("es_MX: claude (high)", "piper/piper-voices/es/es_MX/claude/high/es_MX-claude-high.onnx"),

    # Portuguese (pt_BR/pt_PT)
    ("pt_BR: cadu (medium)", "piper/piper-voices/pt/pt_BR/cadu/medium/pt_BR-cadu-medium.onnx"),
    ("pt_BR: edresson (low)", "piper/piper-voices/pt/pt_BR/edresson/low/pt_BR-edresson-low.onnx"),
    ("pt_BR: faber (medium)", "piper/piper-voices/pt/pt_BR/faber/medium/pt_BR-faber-medium.onnx"),
    ("pt_BR: jeff (medium)", "piper/piper-voices/pt/pt_BR/jeff/medium/pt_BR-jeff-medium.onnx"),
    ("pt_PT: tugao (medium)", "piper/piper-voices/pt/pt_PT/tugao/medium/pt_PT-tugao-medium.onnx"),
]

WORDS_TO_REPLACE = [
    # Hell
    (" The hell ", " The heck "),
    (" The hell! ", " The heck! "),
    (" The hell? ", " The heck? "),
    (" The hell, ", " The heck, "),
    (" The hell. ", " The heck. "),

    (" To hell ", " To heck "),
    (" To hell! ", " To heck! "),
    (" To hell? ", " To heck? "),
    (" To hell, ", " To heck, "),
    (" To hell. ", " To heck. "),

    (" Like hell ", " Like heck "),
    (" Like hell! ", " Like heck! "),
    (" Like hell? ", " Like heck? "),
    (" Like hell, ", " Like heck, "),
    (" Like hell. ", " Like heck. "),

    # Fuck
    (" Motherfuckers ", " Mongrels "),
    (" Motherfuckers! ", " Mongrels! "),
    (" Motherfuckers? ", " Mongrels? "),
    (" Motherfuckers, ", " Mongrels, "),
    (" Motherfuckers. ", " Mongrels. "),

    (" Motherfucker ", " Mongrel "),
    (" Motherfucker! ", " Mongrel! "),
    (" Motherfucker? ", " Mongrel? "),
    (" Motherfucker, ", " Mongrel, "),
    (" Motherfucker. ", " Mongrel. "),

    (" Motherfucking ", " Stupid "),
    (" Motherfucking! ", " Stupid! "),
    (" Motherfucking? ", " Stupid? "),
    (" Motherfucking, ", " Stupid, "),
    (" Motherfucking. ", " Stupid. "),

    (" The fuck ", " The heck "),
    (" The fuck! ", " The heck! "),
    (" The fuck? ", " The heck? "),
    (" The fuck, ", " The heck, "),
    (" The fuck. ", " The heck. "),
    (" The fucking hell ", " The heck "),

    (" Give a fuck ", " Care "),
    (" Give a fuck! ", " Care! "),
    (" Give a fuck? ", " Care? "),
    (" Give a fuck, ", " Care, "),
    (" Give a fuck. ", " Care. "),

    (" Fuck with ", " Mess with "),
    (" Fuck with! ", " Mess with! "),
    (" Fuck with? ", " Mess with? "),
    (" Fuck with, ", " Mess with, "),
    (" Fuck with. ", " Mess with. "),

    (" Fucking with ", " Messing with "),
    (" Fucking with! ", " Messing with! "),
    (" Fucking with? ", " Messing with? "),
    (" Fucking with, ", " Messing with, "),
    (" Fucking with. ", " Messing with. "),

    (" Fucking around ", " Messing around "),
    (" Fucking around! ", " Messing around! "),
    (" Fucking around? ", " Messing around? "),
    (" Fucking around, ", " Messing around, "),
    (" Fucking around. ", " Messing around. "),

    (" Fucked up ", " Messed up "),
    (" Fucked up! ", " Messed up! "),
    (" Fucked up? ", " Messed up? "),
    (" Fucked up, ", " Messed up, "),
    (" Fucked up. ", " Messed up. "),

    (" Fuck up ", " Mess up "),
    (" Fuck up! ", " Mess up! "),
    (" Fuck up? ", " Mess up? "),
    (" Fuck up, ", " Mess up, "),
    (" Fuck up. ", " Mess up. "),

    (" Fuckup ", " Messup "),
    (" Fuckup! ", " Messup! "),
    (" Fuckup? ", " Messup? "),
    (" Fuckup, ", " Messup, "),
    (" Fuckup. ", " Messup. "),

    (" Fuckups ", " Messups "),
    (" Fuckups! ", " Messups! "),
    (" Fuckups? ", " Messups? "),
    (" Fuckups, ", " Messups, "),
    (" Fuckups. ", " Messups. "),

    (" Fucking up ", " Messing up "),
    (" Fucking up! ", " Messing up! "),
    (" Fucking up? ", " Messing up? "),
    (" Fucking up, ", " Messing up, "),
    (" Fucking up. ", " Messing up. "),

    (" Fuckers ", " Mongrels "),
    (" Fuckers! ", " Mongrels! "),
    (" Fuckers? ", " Mongrels? "),
    (" Fuckers, ", " Mongrels, "),
    (" Fuckers. ", " Mongrels. "),

    (" Fucker ", " Mongrel "),
    (" Fucker! ", " Mongrel! "),
    (" Fucker? ", " Mongrel? "),
    (" Fucker, ", " Mongrel, "),
    (" Fucker. ", " Mongrel. "),

    (" As fuck ", " "),
    (" As fuck! ", "! "),
    (" As fuck? ", "? "),
    (" As fuck, ", ", "),
    (" As fuck. ", ". "),

    (" Fuck! ", " Fetch! "),
    (" Fuck, ", " Fetch, "),
    (" Fuck. ", " Fetch. "),
    (" Fuck? ", " Fetch? "),

    # "Fucking" as an intensifier (More complex to catch all cases, but can add some common ones)
    # Note: Replacing just "Fucking" with a word like "Very" or "Really" might change meaning.
    # These replacements aim to slightly soften the intensity.
    (" Fucking good ", " Really good "),
    (" Fucking good! ", " Really good! "),
    (" Fucking good? ", " Really good? "),
    (" Fucking good, ", " Really good, "),
    (" Fucking good. ", " Really good. "),

    (" Fucking great ", " Really great "),
    (" Fucking great! ", " Really great! "),
    (" Fucking great? ", " Really great? "),
    (" Fucking great, ", " Really great, "),
    (" Fucking great. ", " Really great. "),

    (" Fucking impossible ", " Really impossible "),
    (" Fucking impossible! ", " Really impossible! "),
    (" Fucking impossible? ", " Really impossible? "),
    (" Fucking impossible, ", " Really impossible, "),
    (" Fucking impossible. ", " Really impossible. "),

    (" Fucking crazy ", " Really crazy "),
    (" Fucking crazy! ", " Really crazy! "),
    (" Fucking crazy? ", " Really crazy? "),
    (" Fucking crazy, ", " Really crazy, "),
    (" Fucking crazy. ", " Really crazy. "),

    (" Fucking annoying ", " Really annoying "),
    (" Fucking annoying! ", " Really annoying! "),
    (" Fucking annoying? ", " Really annoying? "),
    (" Fucking annoying, ", " Really annoying, "),
    (" Fucking annoying. ", " Really annoying. "),

    (" Fucking amazing ", " Really amazing "),
    (" Fucking amazing! ", " Really amazing! "),
    (" Fucking amazing? ", " Really amazing? "),
    (" Fucking amazing, ", " Really amazing, "),
    (" Fucking amazing. ", " Really amazing. "),

    (" Fucking awesome ", " Really awesome "),
    (" Fucking awesome! ", " Really awesome! "),
    (" Fucking awesome? ", " Really awesome? "),
    (" Fucking awesome, ", " Really awesome, "),
    (" Fucking awesome. ", " Really awesome. "),

    (" Fucking cool ", " Really cool "),
    (" Fucking cool! ", " Really cool! "),
    (" Fucking cool? ", " Really cool? "),
    (" Fucking cool, ", " Really cool, "),
    (" Fucking cool. ", " Really cool. "),

    (" Fucking brilliant ", " Really brilliant "),
    (" Fucking brilliant! ", " Really brilliant! "),
    (" Fucking brilliant? ", " Really brilliant? "),
    (" Fucking brilliant, ", " Really brilliant, "),
    (" Fucking brilliant. ", " Really brilliant. "),

    (" Fucking terrible ", " Really terrible "),
    (" Fucking terrible! ", " Really terrible! "),
    (" Fucking terrible? ", " Really terrible? "),
    (" Fucking terrible, ", " Really terrible, "),
    (" Fucking terrible. ", " Really terrible. "),

    (" Fucking awful ", " Really awful "),
    (" Fucking awful! ", " Really awful! "),
    (" Fucking awful? ", " Really awful? "),
    (" Fucking awful, ", " Really awful, "),
    (" Fucking awful. ", " Really awful. "),

    (" Fucking hard ", " Really hard "),
    (" Fucking hard! ", " Really hard! "),
    (" Fucking hard? ", " Really hard? "),
    (" Fucking hard, ", " Really hard, "),
    (" Fucking hard. ", " Really hard. "),

    (" Fucking difficult ", " Really difficult "),
    (" Fucking difficult! ", " Really difficult! "),
    (" Fucking difficult? ", " Really difficult? "),
    (" Fucking difficult, ", " Really difficult, "),
    (" Fucking difficult. ", " Really difficult. "),

    (" Fucking easy ", " Really easy "),
    (" Fucking easy! ", " Really easy! "),
    (" Fucking easy? ", " Really easy? "),
    (" Fucking easy, ", " Really easy, "),
    (" Fucking easy. ", " Really easy. "),

    # Damn
    (" God damn it ", " Darnit "),
    (" God damn it! ", " Darnit! "),
    (" God damn it? ", " Darnit? "),
    (" God damn it, ", " Darnit, "),
    (" God damn it. ", " Darnit. "),

    (" Goddamn it ", " Darnit "),
    (" Goddamn it! ", " Darnit! "),
    (" Goddamn it? ", " Darnit? "),
    (" Goddamn it, ", " Darnit, "),
    (" Goddamn it. ", " Darnit. "),

    (" God damnit ", " Darnit "),
    (" God damnit! ", " Darnit! "),
    (" God damnit? ", " Darnit? "),
    (" God damnit, ", " Darnit, "),
    (" God damnit. ", " Darnit. "),

    (" Goddamnit ", " Darnit "),
    (" Goddamnit! ", " Darnit! "),
    (" Goddamnit? ", " Darnit? "),
    (" Goddamnit, ", " Darnit, "),
    (" Goddamnit. ", " Darnit. "),

    (" Goddamn ", " Darn "),
    (" Goddamn! ", " Darn! "),
    (" Goddamn? ", " Darn? "),
    (" Goddamn, ", " Darn, "),
    (" Goddamn. ", " Darn. "),

    (" God damn ", " Darn "),
    (" God damn! ", " Darn! "),
    (" God damn? ", " Darn? "),
    (" God damn, ", " Darn, "),
    (" God damn. ", " Darn. "),

    (" Damn it ", " Darnit "),
    (" Damn it! ", " Darnit! "),
    (" Damn it? ", " Darnit? "),
    (" Damn it, ", " Darnit, "),
    (" Damn it. ", " Darnit. "),

    (" Damnit ", " Darnit "),
    (" Damnit! ", " Darnit! "),
    (" Damnit? ", " Darnit? "),
    (" Damnit, ", " Darnit, "),
    (" Damnit. ", " Darnit. "),

    (" Damn! ", " Darn! "),
    (" Damn, ", " Darn, "),
    (" Damn ", " Darn "),
    (" Damn. ", " Darn. "),
    (" Damn? ", " Darn? "),

    (" God damm it ", " Darnit "),
    (" God damm it! ", " Darnit! "),
    (" God damm it? ", " Darnit? "),
    (" God damm it, ", " Darnit, "),
    (" God damm it. ", " Darnit. "),

    (" Goddamm it ", " Darnit "),
    (" Goddamm it! ", " Darnit! "),
    (" Goddamm it? ", " Darnit? "),
    (" Goddamm it, ", " Darnit, "),
    (" Goddamm it. ", " Darnit. "),

    (" God dammit ", " Darnit "),
    (" God dammit! ", " Darnit! "),
    (" God dammit? ", " Darnit? "),
    (" God dammit, ", " Darnit, "),
    (" God dammit. ", " Darnit. "),

    (" Goddammit ", " Darnit "),
    (" Goddammit! ", " Darnit! "),
    (" Goddammit? ", " Darnit? "),
    (" Goddammit, ", " Darnit, "),
    (" Goddammit. ", " Darnit. "),

    (" Goddamm ", " Darn "),
    (" Goddamm! ", " Darn! "),
    (" Goddamm? ", " Darn? "),
    (" Goddamm, ", " Darn, "),
    (" Goddamm. ", " Darn. "),

    (" God damm ", " Darn "),
    (" God damm! ", " Darn! "),
    (" God damm? ", " Darn? "),
    (" God damm, ", " Darn, "),
    (" God damm. ", " Darn. "),

    (" Damm it ", " Darnit "),
    (" Damm it! ", " Darnit! "),
    (" Damm it? ", " Darnit? "),
    (" Damm it, ", " Darnit, "),
    (" Damm it. ", " Darnit. "),

    (" Dammit ", " Darnit "),
    (" Dammit! ", " Darnit! "),
    (" Dammit? ", " Darnit? "),
    (" Dammit, ", " Darnit, "),
    (" Dammit. ", " Darnit. "),

    (" Damm! ", " Darn! "),
    (" Damm, ", " Darn, "),
    (" Damm ", " Darn "),
    (" Damm. ", " Darn. "),
    (" Damm? ", " Darn? "),

    # Shit
    (" Bullshit ", " Nonsense "),
    (" Bullshit! ", " Nonsense! "),
    (" Bullshit? ", " Nonsense? "),
    (" Bullshit, ", " Nonsense, "),
    (" Bullshit. ", " Nonsense. "),

    (" Horsehit ", " Nonsense "),
    (" Horsehit! ", " Nonsense! "),
    (" Horsehit? ", " Nonsense? "),
    (" Horsehit, ", " Nonsense, "),
    (" Horsehit. ", " Nonsense. "),

    (" Shitty ", " Trashy "),
    (" Shitty! ", " Trashy! "),
    (" Shitty? ", " Trashy? "),
    (" Shitty, ", " Trashy, "),
    (" Shitty. ", " Trashy. "),

    (" Shitless ", " Senseless "),
    (" Shitless! ", " Senseless! "),
    (" Shitless? ", " Senseless? "),
    (" Shitless, ", " Senseless, "),
    (" Shitless. ", " Senseless. "),

    (" Cut the shit ", " Cut the nonsense "),
    (" Cut the shit! ", " Cut the nonsense! "),
    (" Cut the shit? ", " Cut the nonsense? "),
    (" Cut the shit, ", " Cut the nonsense, "),
    (" Cut the shit. ", " Cut the nonsense. "),

    (" Piece of shit ", " Piece of trash "),
    (" Piece of shit! ", " Piece of trash! "),
    (" Piece of shit? ", " Piece of trash? "),
    (" Piece of shit, ", " Piece of trash, "),
    (" Piece of shit. ", " Piece of trash. "),

    (" Shithole ", " Trash heap "),
    (" Shithole! ", " Trash heap! "),
    (" Shithole? ", " Trash heap? "),
    (" Shithole, ", " Trash heap, "),
    (" Shithole. ", " Trash heap. "),

    (" Take a shit ", " Take a dump "),
    (" Take a shit! ", " Take a dump! "),
    (" Take a shit? ", " Take a dump? "),
    (" Take a shit, ", " Take a dump, "),
    (" Take a shit. ", " Take a dump. "),

    (" Shit! ", " Crud! "),
    (" Shit? ", " Crud? "),
    (" Shit, ", " Crud, "),
    (" Shit. ", " Crud. "),
    (" Shit ", " Crud "),

    # Crap
    (" Crappy ", " Trashy "),
    (" Crappy! ", " Trashy! "),
    (" Crappy? ", " Trashy? "),
    (" Crappy, ", " Trashy, "),
    (" Crappy. ", " Trashy. "),

    (" Bullcrap ", " Nonsense "),
    (" Bullcrap! ", " Nonsense! "),
    (" Bullcrap? ", " Nonsense? "),
    (" Bullcrap, ", " Nonsense, "),
    (" Bullcrap. ", " Nonsense. "),

    (" Horsecrap ", " Nonsense "),
    (" Horsecrap! ", " Nonsense! "),
    (" Horsecrap? ", " Nonsense? "),
    (" Horsecrap, ", " Nonsense, "),
    (" Horsecrap. ", " Nonsense. "),

    (" Crapton ", " Ton "),
    (" Crapton! ", " Ton! "),
    (" Crapton? ", " Ton? "),
    (" Crapton, ", " Ton, "),
    (" Crapton. ", " Ton. "),

    (" Crap ton ", " Ton "),
    (" Crap ton! ", " Ton! "),
    (" Crap ton? ", " Ton? "),
    (" Crap ton, ", " Ton, "),
    (" Crap ton. ", " Ton. "),

    (" Cut the crap ", " Cut the Nonsense "),
    (" Cut the crap! ", " Cut the Nonsense! "),
    (" Cut the crap? ", " Cut the Nonsense? "),
    (" Cut the crap, ", " Cut the Nonsense, "),
    (" Cut the crap. ", " Cut the Nonsense. "),

    (" Piece of crap ", " Piece of trash "),
    (" Piece of crap! ", " Piece of trash! "),
    (" Piece of crap? ", " Piece of trash? "),
    (" Piece of crap, ", " Piece of trash, "),
    (" Piece of crap. ", " Piece of trash. "),

    (" Craphole ", " Trash heap "),
    (" Craphole! ", " Trash heap! "),
    (" Craphole? ", " Trash heap? "),
    (" Craphole, ", " Trash heap, "),
    (" Craphole. ", " Trash heap. "),

    (" Crapped ", " Pooped "),
    (" Take a crap ", " Take a dump "),
    (" Take a crap! ", " Take a dump! "),
    (" Take a crap? ", " Take a dump? "),
    (" Take a crap, ", " Take a dump, "),
    (" Take a crap. ", " Take a dump. "),

    (" Crap! ", " Crud! "),
    (" Crap? ", " Crud? "),
    (" Crap, ", " Crud, "),
    (" Crap. ", " Crud. "),
    (" Crap ", " Crud "),

    # Bastard
    (" Bastards ", " Mongrels "),
    (" Bastards! ", " Mongrels! "),
    (" Bastards? ", " Mongrels? "),
    (" Bastards, ", " Mongrels, "),
    (" Bastards. ", " Mongrels. "),

    (" Bastard ", " Mongrel "),
    (" Bastard! ", " Mongrel! "),
    (" Bastard? ", " Mongrel? "),
    (" Bastard, ", " Mongrel, "),
    (" Bastard. ", " Mongrel. "),

    # Retard
    (" Retarded ", " Idiotic "),
    (" Retarded! ", " Idiotic! "),
    (" Retarded? ", " Idiotic? "),
    (" Retarded, ", " Idiotic, "),
    (" Retarded. ", " Idiotic. "),

    (" Retards ", " Idiots "),
    (" Retards! ", " Idiots! "),
    (" Retards? ", " Idiots? "),
    (" Retards, ", " Idiots, "),
    (" Retards. ", " Idiots. "),

    (" Retard ", " Idiot "),
    (" Retard! ", " Idiot! "),
    (" Retard? ", " Idiot? "),
    (" Retard, ", " Idiot, "),
    (" Retard. ", " Idiot. "),

    # Bitch
    (" Bitchy ", " Annoying "),
    (" Bitchy! ", " Annoying! "),
    (" Bitchy? ", " Annoying? "),
    (" Bitchy, ", " Annoying, "),
    (" Bitchy. ", " Annoying. "),

    (" Bitching ", " Whining "),
    (" Bitching! ", " Whining! "),
    (" Bitching? ", " Whining? "),
    (" Bitching, ", " Whining, "),
    (" Bitching. ", " Whining. "),

    (" Bitches! ", " Shrews! "),
    (" Bitches? ", " Shrews? "),
    (" Bitches, ", " Shrews, "),
    (" Bitches. ", " Shrews. "),
    (" Bitches ", " Shrews "),

    (" Son of a bitch ", " Scoundrel "),
    (" Son of a bitch! ", " Scoundrel! "),
    (" Son of a bitch? ", " Scoundrel? "),
    (" Son of a bitch, ", " Scoundrel, "),
    (" Son of a bitch. ", " Scoundrel. "),

    (" Sonofabitch ", " Scoundrel "),
    (" Sonofabitch! ", " Scoundrel! "),
    (" Sonofabitch? ", " Scoundrel? "),
    (" Sonofabitch, ", " Scoundrel, "),
    (" Sonofabitch. ", " Scoundrel. "),

    (" Bitch! ", " Shrew! "),
    (" Bitch? ", " Shrew? "),
    (" Bitch, ", " Shrew, "),
    (" Bitch. ", " Shrew. "),
    (" Bitch ", " Shrew "),

    # Ass
    (" Assholes ", " Jerks "),
    (" Assholes! ", " Jerks! "),
    (" Assholes? ", " Jerks? "),
    (" Assholes, ", " Jerks, "),
    (" Assholes. ", " Jerks. "),

    (" Asshole ", " Jerk "),
    (" Asshole! ", " Jerk! "),
    (" Asshole? ", " Jerk? "),
    (" Asshole, ", " Jerk, "),
    (" Asshole. ", " Jerk. "),

    (" Dumbasses ", " Morons "),
    (" Dumbasses! ", " Morons! "),
    (" Dumbasses? ", " Morons? "),
    (" Dumbasses, ", " Morons, "),
    (" Dumbasses. ", " Morons. "),

    (" Dumbass ", " Moron "),
    (" Dumbass! ", " Moron! "),
    (" Dumbass? ", " Moron? "),
    (" Dumbass, ", " Moron, "),
    (" Dumbass. ", " Moron. "),

    (" Pain in the ass ", " Pain in the rear "),
    (" Pain in the ass! ", " Pain in the rear! "),
    (" Pain in the ass? ", " Pain in the rear? "),
    (" Pain in the ass, ", " Pain in the rear, "),
    (" Pain in the ass. ", " Pain in the rear. "),

    (" Such an ass ", " Such a jerk "),
    (" Such an ass! ", " Such a jerk! "),
    (" Such an ass? ", " Such a jerk? "),
    (" Such an ass, ", " Such a jerk, "),
    (" Such an ass. ", " Such a jerk. "),

    (" Up his ass ", " Up his rear "),
    (" Up his ass! ", " Up his rear! "),
    (" Up his ass? ", " Up his rear? "),
    (" Up his ass, ", " Up his rear, "),
    (" Up his ass. ", " Up his rear. "),
    (" Up her ass ", " Up her rear "),
    (" Up her ass! ", " Up her rear! "),
    (" Up her ass? ", " Up her rear? "),
    (" Up her ass, ", " Up her rear, "),
    (" Up her ass. ", " Up her rear. "),
    (" Up their ass ", " Up their rear "),
    (" Up their ass! ", " Up their rear! "),
    (" Up their ass? ", " Up their rear? "),
    (" Up their ass, ", " Up their rear, "),
    (" Up their ass. ", " Up their rear. "),
    (" Up our ass ", " Up our rear "),
    (" Up our ass! ", " Up our rear! "),
    (" Up our ass? ", " Up our rear? "),
    (" Up our ass, ", " Up our rear, "),
    (" Up our ass. ", " Up our rear. "),
    (" Up its ass ", " Up its rear "),
    (" Up its ass! ", " Up its rear! "),
    (" Up its ass? ", " Up its rear? "),
    (" Up its ass, ", " Up its rear, "),
    (" Up its ass. ", " Up its rear. "),
    (" Up my ass ", " Up my rear "),
    (" Up my ass! ", " Up my rear! "),
    (" Up my ass? ", " Up my rear? "),
    (" Up my ass, ", " Up my rear, "),
    (" Up my ass. ", " Up my rear. "),

    # Kick/Ass
    (" Kick ass! ", " Kick butt! "),
    (" Kick ass? ", " Kick butt? "),
    (" Kick ass, ", " Kick butt, "),
    (" Kick ass. ", " Kick butt. "),
    (" Kick ass ", " Kick butt "),

    (" Kickass! ", " Awesome! "),
    (" Kickass? ", " Awesome? "),
    (" Kickass, ", " Awesome, "),
    (" Kickass. ", " Awesome. "),
    (" Kickass ", " Awesome "),

    (" Kick my ass! ", " Kick my butt! "),
    (" Kick my ass? ", " Kick my butt? "),
    (" Kick my ass, ", " Kick my butt, "),
    (" Kick my ass. ", " Kick my butt. "),
    (" Kick my ass ", " Kick my butt "),
    (" Kick your ass! ", " Kick your butt! "),
    (" Kick your ass? ", " Kick your butt? "),
    (" Kick your ass, ", " Kick your butt, "),
    (" Kick your ass. ", " Kick your butt. "),
    (" Kick your ass ", " Kick your butt "),
    (" Kick his ass! ", " Kick his butt! "),
    (" Kick his ass? ", " Kick his butt? "),
    (" Kick his ass, ", " Kick his butt, "),
    (" Kick his ass. ", " Kick his butt. "),
    (" Kick his ass ", " Kick his butt "),
    (" Kick her ass! ", " Kick her butt! "),
    (" Kick her ass? ", " Kick her butt? "),
    (" Kick her ass, ", " Kick her butt, "),
    (" Kick her ass. ", " Kick her butt. "),
    (" Kick her ass ", " Kick her butt "),
    (" Kick their ass! ", " Kick their butt! "),
    (" Kick their ass? ", " Kick their butt? "),
    (" Kick their ass, ", " Kick their butt, "),
    (" Kick their ass. ", " Kick their butt. "),
    (" Kick their ass ", " Kick their butt "),
    (" Kick our ass! ", " Kick our butt! "),
    (" Kick our ass? ", " Kick our butt? "),
    (" Kick our ass, ", " Kick our butt, "),
    (" Kick our ass. ", " Kick our butt. "),
    (" Kick our ass ", " Kick our butt "),
    (" Kick its ass! ", " Kick its butt! "),
    (" Kick its ass? ", " Kick its butt? "),
    (" Kick its ass, ", " Kick its butt, "),
    (" Kick its ass. ", " Kick its butt. "),
    (" Kick its ass ", " Kick its butt "),

    # Kiss/Ass
    (" Kissass! ", " Sycophant! "),
    (" Kissass? ", " Sycophant? "),
    (" Kissass, ", " Sycophant, "),
    (" Kissass. ", " Sycophant. "),
    (" Kissass ", " Sycophant "),

    (" Kiss ass! ", " Flatter! "),
    (" Kiss ass? ", " Flatter? "),
    (" Kiss ass, ", " Flatter, "),
    (" Kiss ass. ", " Flatter. "),
    (" Kiss ass ", " Flatter "),

    (" Kiss my ass! ", " Flatter me! "),
    (" Kiss my ass? ", " Flatter me? "),
    (" Kiss my ass, ", " Flatter me, "),
    (" Kiss my ass. ", " Flatter me. "),
    (" Kiss my ass ", " Flatter me "),
    (" Kiss your ass! ", " Flatter you! "),
    (" Kiss your ass? ", " Flatter you? "),
    (" Kiss your ass, ", " Flatter you, "),
    (" Kiss your ass. ", " Flatter you. "),
    (" Kiss your ass ", " Flatter "),
    (" Kiss his ass! ", " Flatter him! "),
    (" Kiss his ass? ", " Flatter him? "),
    (" Kiss his ass, ", " Flatter him, "),
    (" Kiss his ass. ", " Flatter him. "),
    (" Kiss his ass ", " Flatter him "),
    (" Kiss her ass! ", " Flatter her! "),
    (" Kiss her ass? ", " Flatter her? "),
    (" Kiss her ass, ", " Flatter her, "),
    (" Kiss her ass. ", " Flatter her. "),
    (" Kiss her ass ", " Flatter her "),
    (" Kiss their ass! ", " Flatter them! "),
    (" Kiss their ass? ", " Flatter them? "),
    (" Kiss their ass, ", " Flatter them, "),
    (" Kiss their ass. ", " Flatter them. "),
    (" Kiss their ass ", " Flatter them "),
    (" Kiss our ass! ", " Flatter us! "),
    (" Kiss our ass? ", " Flatter us? "),
    (" Kiss our ass, ", " Flatter us, "),
    (" Kiss our ass. ", " Flatter us. "),
    (" Kiss our ass ", " Flatter us "),
    (" Kiss its ass! ", " Flatter it! "),
    (" Kiss its ass? ", " Flatter it? "),
    (" Kiss its ass, ", " Flatter it, "),
    (" Kiss its ass. ", " Flatter it. "),
    (" Kiss its ass ", " Flatter it "),

    # Piss
    (" Take a piss ", " Take a pee "),
    (" Take a piss! ", " Take a pee! "),
    (" Take a piss? ", " Take a pee? "),
    (" Take a piss, ", " Take a pee, "),
    (" Take a piss. ", " Take a pee. "),
    (" Pissed off ", " Ticked off "),
    (" Pissed off! ", " Ticked off! "),
    (" Pissed off? ", " Ticked off? "),
    (" Pissed off, ", " Ticked off, "),
    (" Pissed off. ", " Ticked off. "),

    # Dick
    (" Such a dick ", " Such a jerk "),
    (" Such a dick! ", " Such a jerk! "),
    (" Such a dick? ", " Such a jerk? "),
    (" Such a dick, ", " Such a jerk, "),
    (" Such a dick. ", " Such a jerk. "),
    (" Be a dick ", " Be a jerk "),
    (" Be a dick! ", " Be a jerk! "),
    (" Be a dick? ", " Be a jerk? "),
    (" Be a dick, ", " Be a jerk, "),
    (" Be a dick. ", " Be a jerk. "),
    (" Being a dick ", " Being a jerk "),
    (" Being a dick! ", " Being a jerk! "),
    (" Being a dick? ", " Being a jerk? "),
    (" Being a dick, ", " Being a jerk, "),
    (" Being a dick. ", " Being a jerk. "),

    # Jesus/Christ
    (" Jesus! ", " Fetch! "),
    (" Jesus? ", " Fetch? "),
    (" Jesus, ", " Fetch, "),
    (" Jesus. ", " Fetch. "),
    (" Jesus ", " Fetch "),
    (" Christ! ", " Fetch! "),
    (" Christ? ", " Fetch? "),
    (" Christ, ", " Fetch, "),
    (" Christ. ", " Fetch. "),
    (" Christ ", " Fetch "),
    (" Jesus Christ! ", " Fetch! "),
    (" Jesus Christ? ", " Fetch? "),
    (" Jesus Christ, ", " Fetch, "),
    (" Jesus Christ. ", " Fetch. "),
    (" Jesus Christ ", " Fetch "),
    (" For Jesus's sake! ", " For Pete's sake! "),
    (" For Jesus's sake? ", " For Pete's sake? "),
    (" For Jesus's sake, ", " For Pete's sake, "),
    (" For Jesus's sake. ", " For Pete's sake. "),
    (" For Jesus's sake ", " For Pete's sake "),
    (" For Christ's sake! ", " For Pete's sake! "),
    (" For Christ's sake? ", " For Pete's sake? "),
    (" For Christ's sake, ", " For Pete's sake, "),
    (" For Christ's sake. ", " For Pete's sake. "),
    (" For Christ's sake ", " For Pete's sake "),
    (" For God's sake! ", " For Pete's sake! "),
    (" For God's sake? ", " For Pete's sake? "),
    (" For God's sake, ", " For Pete's sake, "),
    (" For God's sake. ", " For Pete's sake. "),
    (" For God's sake ", " For Pete's sake "),

    # --- Additions Start Here ---

    # Cunt (Note: This is a very strong word, replacements aim for strong disapproval)
    (" Cunts ", " Vile individuals "),
    (" Cunts! ", " Vile individuals! "),
    (" Cunts? ", " Vile individuals? "),
    (" Cunts, ", " Vile individuals, "),
    (" Cunts. ", " Vile individuals. "),

    (" Cunt ", " Vile individual "),
    (" Cunt! ", " Vile individual! "),
    (" Cunt? ", " Vile individual? "),
    (" Cunt, ", " Vile individual, "),
    (" Cunt. ", " Vile individual. "),

    # Cock (Referring to insult, not anatomy)
    (" Such a cock ", " Such a fool "),
    (" Such a cock! ", " Such a fool! "),
    (" Such a cock? ", " Such a fool? "),
    (" Such a cock, ", " Such a fool, "),
    (" Such a cock. ", " Such a fool. "),

    (" Be a cock ", " Be a fool "),
    (" Be a cock! ", " Be a fool! "),
    (" Be a cock? ", " Be a fool? "),
    (" Be a cock, ", " Be a fool, "),
    (" Be a cock. ", " Be a fool. "),

    # Prick (Similar to Dick/Cock as an insult)
    (" Such a prick ", " Such a nuisance "),
    (" Such a prick! ", " Such a nuisance! "),
    (" Such a prick? ", " Such a nuisance? "),
    (" Such a prick, ", " Such a nuisance, "),
    (" Such a prick. ", " Such a nuisance. "),

    (" Be a prick ", " Be a nuisance "),
    (" Be a prick! ", " Be a nuisance! "),
    (" Be a prick? ", " Be a nuisance? "),
    (" Be a prick, ", " Be a nuisance, "),
    (" Be a prick. ", " Be a nuisance. "),

    # Arse / Arsehole (British variants of Ass/Asshole)
    (" Arseholes ", " Jerks "), # Using the same replacement as Assholes
    (" Arseholes! ", " Jerks! "),
    (" Arseholes? ", " Jerks? "),
    (" Arseholes, ", " Jerks, "),
    (" Arseholes. ", " Jerks. "),

    (" Arsehole ", " Jerk "), # Using the same replacement as Asshole
    (" Arsehole! ", " Jerk! "),
    (" Arsehole? ", " Jerk? "),
    (" Arsehole, ", " Jerk, "),
    (" Arsehole. ", " Jerk. "),

    (" Pain in the arse ", " Pain in the rear "), # Using the same replacement as Pain in the ass
    (" Pain in the arse! ", " Pain in the rear! "),
    (" Pain in the arse? ", " Pain in the rear? "),
    (" Pain in the arse, ", " Pain in the rear, "),
    (" Pain in the arse. ", " Pain in the rear. "),

    (" Such an arse ", " Such a jerk "), # Using the same replacement as Such an ass
    (" Such an arse! ", " Such a jerk! "),
    (" Such an arse? ", " Such a jerk? "),
    (" Such an arse, ", " Such a jerk, "),
    (" Such an arse. ", " Such a jerk. "),

    (" Arse! ", " Rear! "),
    (" Arse? ", " Rear? "),
    (" Arse, ", " Rear, "),
    (" Arse. ", " Rear. "),
    (" Arse ", " Rear "),

    # Wanker (British insult)
    (" Wankers ", " Buffoons "),
    (" Wankers! ", " Buffoons! "),
    (" Wankers? ", " Buffoons? "),
    (" Wankers, ", " Buffoons, "),
    (" Wankers. ", " Buffoons. "),

    (" Wanker ", " Buffoon "),
    (" Wanker! ", " Buffoon! "),
    (" Wanker? ", " Buffoon? "),
    (" Wanker, ", " Buffoon, "),
    (" Wanker. ", " Buffoon. "),

    # Tosser (British insult)
    (" Tossers ", " Simpletons "),
    (" Tossers! ", " Simpletons! "),
    (" Tossers? ", " Simpletons? "),
    (" Tossers, ", " Simpletons, "),
    (" Tossers. ", " Simpletons. "),

    (" Tosser ", " Simpleton "),
    (" Tosser! ", " Simpleton! "),
    (" Tosser? ", " Simpleton? "),
    (" Tosser, ", " Simpleton, "),
    (" Tosser. ", " Simpleton. "),

    # Bollocks (British exclamation/nonsense)
    (" Bollocks! ", " Poppycock! "),
    (" Bollocks? ", " Poppycock? "),
    (" Bollocks, ", " Poppycock, "),
    (" Bollocks. ", " Poppycock. "),
    (" Bollocks ", " Poppycock "),

    # Bloody (British intensifier/mild swear)
    (" Bloody hell ", " Heck "), # You have "The fucking hell", adding this British variant
    (" Bloody hell! ", " Heck! "),
    (" Bloody hell? ", " Heck? "),
    (" Bloody hell, ", " Heck, "),
    (" Bloody hell. ", " Heck. "),

    # Bugger (British mild swear)
    (" Bugger off ", " Go away "),
    (" Bugger off! ", " Go away! "),
    (" Bugger off? ", " Go away? "),
    (" Bugger off, ", " Go away, "),
    (" Bugger off. ", " Go away. "),

    (" Bugger! ", " Bother! "),
    (" Bugger? ", " Bother? "),
    (" Bugger, ", " Bother, "),
    (" Bugger. ", " Bother. "),
    (" Bugger ", " Bother "),

    # Whore
    (" Whores ", " Harlots "),
    (" Whores! ", " Harlots! "),
    (" Whores? ", " Harlots? "),
    (" Whores, ", " Harlots, "),
    (" Whores. ", " Harlots. "),

    (" Whore ", " Harlot "),
    (" Whore! ", " Harlot! "),
    (" Whore? ", " Harlot? "),
    (" Whore, ", " Harlot, "),
    (" Whore. ", " Harlot. "),

    # Slut
    (" Sluts ", " Hussies "),
    (" Sluts! ", " Hussies! "),
    (" Sluts? ", " Hussies? "),
    (" Sluts, ", " Hussies, "),
    (" Sluts. ", " Hussies. "),

    (" Slut ", " Hussy "),
    (" Slut! ", " Hussy! "),
    (" Slut? ", " Hussy? "),
    (" Slut, ", " Hussy, "),
    (" Slut. ", " Hussy. "),

    # Douchebag / Douche
    (" Douchebags ", " Louts "),
    (" Douchebags! ", " Louts! "),
    (" Douchebags? ", " Louts? "),
    (" Douchebags, ", " Louts, "),
    (" Douchebags. ", " Louts. "),

    (" Douchebag ", " Lout "),
    (" Douchebag! ", " Lout! "),
    (" Douchebag? ", " Lout? "),
    (" Douchebag, ", " Lout, "),
    (" Douchebag. ", " Lout. "),

    (" Douche! ", " Oaf! "),
    (" Douche? ", " Oaf? "),
    (" Douche, ", " Oaf, "),
    (" Douche. ", " Oaf. "),
    (" Douche ", " Oaf "),

    # Idiot / Moron (As standalone words)
    (" Idiots ", " Simpletons "), # Using the same replacement as Tossers
    (" Idiots! ", " Simpletons! "),
    (" Idiots? ", " Simpletons? "),
    (" Idiots, ", " Simpletons, "),
    (" Idiots. ", " Simpletons. "),

    (" Idiot ", " Simpleton "), # Using the same replacement as Tosser
    (" Idiot! ", " Simpleton! "),
    (" Idiot? ", " Simpleton? "),
    (" Idiot, ", " Simpleton, "),
    (" Idiot. ", " Simpleton. "),

    (" Morons ", " Dimwits "),
    (" Morons! ", " Dimwits! "),
    (" Morons? ", " Dimwits? "),
    (" Morons, ", " Dimwits, "),
    (" Morons. ", " Dimwits. "),

    (" Moron ", " Dimwit "),
    (" Moron! ", " Dimwit! "),
    (" Moron? ", " Dimwit? "),
    (" Moron, ", " Dimwit, "),
    (" Moron. ", " Dimwit. "),

    # Compound insults (Examples - this category is vast)
    (" Shit-faced ", " Drunk "),
    (" Shit-faced! ", " Drunk! "),
    (" Shit-faced? ", " Drunk? "),
    (" Shit-faced, ", " Drunk, "),
    (" Shit-faced. ", " Drunk. "),

    (" Ass-backward ", " Confused "),
    (" Ass-backward! ", " Confused! "),
    (" Ass-backward? ", " Confused? "),
    (" Ass-backward, ", " Confused, "),
    (" Ass-backward. ", " Confused. "),
]

# Moved get_project_root outside the class if it's logically tied to the GUI instance
# If it's a utility, it could remain outside. Let's keep it outside for now
# as it doesn't rely on instance state.
def get_project_root():
    """Finds the project root directory assuming it's named 'glowing-umbrella'."""
    path = os.path.abspath(os.getcwd())
    while True:
        if os.path.basename(path).lower() == "glowing-umbrella": # Use lower() for robustness
            return path
        parent = os.path.dirname(path)
        if parent == path:
            # Fallback: If not in a specific project structure, use current directory
            print("Warning: Could not find 'glowing-umbrella' project root. Using current directory.")
            return os.getcwd()
        path = parent


class PiperTTSGUI:
    def __init__(self, root):
        self.root = root
        root.title("Glowing Umbrella - Piper TTS Configuration")

        # Instance variables to hold file paths using StringVars for easy updating in Entry widgets
        self.input_folder_path = tk.StringVar() # For Text File Conversion
        self.output_folder_path = tk.StringVar() # For Text File Conversion Output
        self.text_file_path = tk.StringVar() # For Text File Split/Replace

        self.create_widgets()
        self.setup_validation() # Moved setup_validation call here

    def create_widgets(self):
        # Use a frame for better layout management if needed, but grid is fine for this
        # main_frame = ttk.Frame(self.root, padding="10")
        # main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # --- Text File Conversion Section ---
        ttk.Label(self.root, text="Input Text Folder:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.input_path_entry = ttk.Entry(self.root, width=50, textvariable=self.input_folder_path)
        self.input_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.root, text="Browse...", command=self.browse_input).grid(row=0, column=2, padx=5, sticky="w")

        ttk.Label(self.root, text="Output Audio Folder:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.output_path_entry = ttk.Entry(self.root, width=50, textvariable=self.output_folder_path)
        self.output_path_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.root, text="Browse...", command=self.browse_output).grid(row=1, column=2, padx=5, sticky="w")

        # Voice
        ttk.Label(self.root, text="Voice:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.voice_combo = ttk.Combobox(
            self.root,
            values=[v[0] for v in VOICES],  # Combine lang tag and name
            state="readonly" # Prevent typing in the combobox
        )
        self.voice_combo.current(1) # Set default selected voice
        self.voice_combo.grid(row=2, column=1, padx=5, pady=5, columnspan=2, sticky="ew") # Span 2 columns

        # Speed/Pause
        ttk.Label(self.root, text="Speed Scale (0.001-2.0):").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.speed_entry = ttk.Entry(self.root)
        self.speed_entry.insert(0, "0.9") # Default value
        self.speed_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(self.root, text="Pause (0.001-2.0):").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.pause_entry = ttk.Entry(self.root)
        self.pause_entry.insert(0, "0.2") # Default value
        self.pause_entry.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

        # Convert Button for Text Files
        ttk.Button(self.root, text="Convert Text Files to Audio", command=self.convert_files).grid(row=5, column=0, columnspan=3, pady=10)

        # Add a separator
        ttk.Separator(self.root, orient='horizontal').grid(row=6, column=0, columnspan=3, sticky="ew", pady=10)


        # --- Text File Utilities Section ---
        ttk.Label(self.root, text="Text File Utilities:").grid(row=7, column=0, padx=5, pady=5, sticky="w", columnspan=3)

        ttk.Label(self.root, text="Text File:").grid(row=8, column=0, padx=5, pady=5, sticky="w")
        self.txt_path_entry = ttk.Entry(self.root, width=50, textvariable=self.text_file_path)
        self.txt_path_entry.grid(row=8, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.root, text="Browse...", command=self.browse_txt).grid(row=8, column=2, padx=5, sticky="w")

        ttk.Label(self.root, text="Words per Split Text File:").grid(row=9, column=0, padx=5, pady=5, sticky="w")
        self.words_entry = ttk.Entry(self.root)
        self.words_entry.insert(0, "3700") # Default value
        self.words_entry.grid(row=9, column=1, padx=5, pady=5, sticky='ew')

        ttk.Button(self.root, text="Split Text File", command=self.split_text_file).grid(row=10, column=0, pady=10)
        ttk.Button(self.root, text="Replace Words in Text File", command=self.replace_words_in_text_file).grid(row=10, column=1, pady=10)


        # Add a separator
        ttk.Separator(self.root, orient='horizontal').grid(row=11, column=0, columnspan=3, sticky="ew", pady=10)


        # --- Other Utilities ---
        # Adjusted row number after removing Word section
        ttk.Button(self.root, text="Generate Voice Samples", command=self.generate_samples).grid(row=12, column=0, columnspan=3, pady=10)


        # Configure column weights so the entry fields expand
        self.root.grid_columnconfigure(1, weight=1)


    def setup_validation(self):
        """Sets up input validation for relevant entry fields."""
        # Register validation commands once
        vcmd_float = self.root.register(self.validate_float)
        vcmd_int = self.root.register(self.validate_int)

        # Apply validation to the entry fields
        self.speed_entry.config(validate="key", validatecommand=(vcmd_float, '%P'))
        self.pause_entry.config(validate="key", validatecommand=(vcmd_float, '%P'))
        self.words_entry.config(validate="key", validatecommand=(vcmd_int, '%P'))


    def validate_float(self, value):
        """Validates if the input value is a valid float string format or empty."""
        # Allow empty string (for deleting)
        if not value:
            return True

        # Check if the string can be potentially converted to a float
        try:
            float(value)
            # If float conversion succeeds, the format is okay for 'key' validation
            return True
        except ValueError:
            # If float conversion fails, the format is invalid (e.g., "abc", "1.2.3")
            return False


    def validate_int(self, value):
        """Validates if the input value contains only digits or is empty."""
        # Allow empty string (for deleting)
        if not value:
            return True

        # Check if the string consists only of digits
        # Use str.isdigit() which is simple and sufficient here
        return value.isdigit()


    # --- Browse methods ---
    def browse_input(self):
        """Opens a dialog to select an input folder for text files."""
        path = filedialog.askdirectory(title="Select Input Folder (Text Files)")
        if path:
            self.input_folder_path.set(path)


    def browse_output(self):
        """Opens a dialog to select an output folder for audio files."""
        path = filedialog.askdirectory(title="Select Output Folder (Audio Files)")
        if path:
            self.output_folder_path.set(path)

    def browse_txt(self):
        """Opens a dialog to select a single text file for utilities."""
        path = filedialog.askopenfilename(
            title="Select Text File",
            filetypes=[("Text files", "*.txt"), ("All Files", "*.*")]
        )
        if path:
            self.text_file_path.set(path)


    # --- Text File Conversion Method ---
    def convert_files(self):
        """Converts text files from input folder to audio using Piper TTS."""
        try:
            # Get selected voice/model
            model_idx = self.voice_combo.current()
            if model_idx == -1:
                messagebox.showerror("Error", "Select a voice first.")
                return
            model_relpath = VOICES[model_idx][1]

            input_dir = self.input_folder_path.get()
            output_dir = self.output_folder_path.get()

            if not os.path.isdir(input_dir):
                messagebox.showerror("Error", f"Input directory not found:\n{input_dir}")
                return
            if not output_dir:
                messagebox.showerror("Error", "Output directory is not selected.")
                return
            # Ensure output directory exists
            os.makedirs(output_dir, exist_ok=True)

            # --- Get and Validate Speed/Pause Values ---
            try:
                length_scale_str = self.speed_entry.get()
                sentence_silence_str = self.pause_entry.get()

                # Perform float conversion and range check here, where the value is used
                length_scale = float(length_scale_str)
                sentence_silence = float(sentence_silence_str)

                if not (0.001 <= length_scale <= 2.0):
                     messagebox.showerror("Error", f"Speed Scale value {length_scale_str} is outside the valid range (0.001 to 2.0).")
                     return
                if not (0.001 <= sentence_silence <= 2.0):
                     messagebox.showerror("Error", f"Pause value {sentence_silence_str} is outside the valid range (0.001 to 2.0).")
                     return

            except ValueError:
                # This might happen if the user somehow bypassed key validation or the field is empty
                 messagebox.showerror("Error", "Invalid speed or pause value format. Please enter valid numbers.")
                 return
            # -------------------------------------------


            # Find project root
            project_root = get_project_root() # Call the function

            # Build absolute model path
            model_path_abs = os.path.join(project_root, model_relpath) # Use abs var name
            if not os.path.isfile(model_path_abs):
                messagebox.showerror("Error", f"Model file not found:\n{model_path_abs}")
                return

            # Check for piper executable
            piper_executable_abs = os.path.join(project_root, "piper_win", "piper.exe") # Use abs var name
            if not os.path.isfile(piper_executable_abs):
                messagebox.showerror("Error", f"Piper executable not found:\n{piper_executable_abs}")
                return

            # Process each .txt file in input_dir
            files_converted = 0
            processed_files = 0
            # Use listdir and filter first to get total count for potentially better progress reporting later
            txt_files = [f for f in os.listdir(input_dir) if f.lower().endswith(".txt")]
            total_files_to_process = len(txt_files)

            if total_files_to_process == 0:
                messagebox.showinfo("Info", f"No .txt files found in input directory:\n{input_dir}")
                return # Exit early if no files

            print(f"Starting conversion of {total_files_to_process} text file(s) from {input_dir}...")
            print("-" * 40) # Separator

            for filename in txt_files:
                processed_files += 1
                base = os.path.splitext(filename)[0]
                input_path_file = os.path.join(input_dir, filename) # Full path to the input .txt file
                output_path_file = os.path.join(output_dir, f"{base}.wav") # Full path to the output .wav file

                # --- Pre-process paths for the bash command (convert backslashes to forward slashes) ---
                input_path_bash = input_path_file.replace("\\", "/")
                output_path_bash = output_path_file.replace("\\", "/")
                piper_executable_bash = piper_executable_abs.replace("\\", "/") # Use the absolute path
                model_path_bash = model_path_abs.replace("\\", "/") # Use the absolute path
                # -------------------------------------------------------------------------------------

                # Print start processing message with count
                print(f"[{processed_files}/{total_files_to_process}] Processing: {filename}")

                # Construct the command using bash -c for robust path handling
                command = (
                    f'bash -c "MSYS2_ARG_CONV_EXCL=\\"*\\" cat \\"{input_path_bash}\\" | '
                    f'\\"{piper_executable_bash}\\" '
                    f'--model \\"{model_path_bash}\\" --length_scale {length_scale:.3f} --sentence_silence {sentence_silence:.3f} ' # Use :.3f for precision
                    f'--output_file \\"{output_path_bash}\\""'
                )

                # print("COMMAND:", command) # Uncomment for debugging the command string
                result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='utf-8')

                # Piper TTS often prints progress/info to stderr, which might look like errors but isn't necessarily.
                # If result.returncode is 0, consider stderr as informational output from Piper.
                if result.stderr and result.returncode == 0:
                    print(f"  Info from Piper for {filename}:\n{result.stderr.strip()}") # Print stderr if successful (often progress)

                if result.returncode == 0:
                    files_converted += 1
                    print(f"  --> Successfully converted: {filename} -> {os.path.basename(output_path_file)}") # Success message
                else:
                    print(f"  --> FAILED to convert: {filename}") # Failure message
                    # Print command, return code, and stderr on failure for debugging
                    print(f"    Command: {command}")
                    print(f"    Return Code: {result.returncode}")
                    print(f"    Stderr: {result.stderr.strip()}")
                    # Show warning message, but allow processing to continue for other files
                    messagebox.showwarning(f"Conversion Failed for {filename}", f"Piper command failed with return code {result.returncode}.\nDetails in console.")

                print("-" * 40) # Separator after each file

            # Final summary messagebox
            messagebox.showinfo("Done", f"Attempted to convert {processed_files} files.\nSuccessfully converted {files_converted} files to audio.")
            print("Conversion process finished.")

        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {e}")
            print(f"An unexpected error occurred during conversion: {e}")

    # --- Text File Utility Methods ---
    def split_text_file(self):
        """Splits a text file by words."""
        txt_path = self.text_file_path.get()
        if not os.path.isfile(txt_path):
            messagebox.showerror("Error", "Selected text file does not exist.")
            return

        # --- Get and Validate Words per File ---
        try:
            words_per_file_str = self.words_entry.get()
            # Perform int conversion and range check here
            words_per_file = int(words_per_file_str)

            if words_per_file < 1 or words_per_file > 100000: # Validate range
                messagebox.showerror("Error", f"Words per file value {words_per_file_str} is outside the valid range (1 to 100000).")
                return

        except ValueError:
             # This might happen if the user somehow bypassed key validation or the field is empty
            messagebox.showerror("Error", "Invalid value for words per file. Please enter a whole number.")
            return
        # -------------------------------------


        try:
            # Note: Split Text File currently only tries UTF-8. If you get decode errors here,
            # you might need to add the multi-encoding logic from replace_words_in_text_file.
            with open(txt_path, "r", encoding="utf-8") as f:
                words = f.read().split()
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file:\n{e}")
            return

        if not words:
            messagebox.showinfo("Info", "The text file is empty or contains only whitespace.")
            return


        dir_name = os.path.dirname(txt_path)
        base_name = os.path.basename(txt_path)
        base, ext = os.path.splitext(base_name)
        if not ext: ext = ".txt" # Default to .txt if no extension


        file_idx = 1 # Start indexing from 1
        total_words = len(words)
        files_created = 0

        try:
            for i in range(0, total_words, words_per_file):
                chunk_words = words[i:i+words_per_file]
                # Generate output filename like original_part1.txt
                out_filename = f"{base}_part{file_idx}{ext}" # Changed to 'part' and start index from 1
                out_path = os.path.join(dir_name, out_filename)

                # Simple overwrite prevention
                counter = 1
                original_out_name_for_msg = out_filename # Use this for user message if conflict occurs
                while os.path.exists(out_path):
                    out_filename = f"{base}_part{file_idx}_{counter}{ext}" # Adjusted naming for conflict
                    out_path = os.path.join(dir_name, out_filename)
                    counter += 1
                if counter > 1:
                    print(f"Warning: File '{original_out_name_for_msg}' existed, saving as '{os.path.basename(out_path)}'")


                with open(out_path, "w", encoding="utf-8") as out_file:
                    out_file.write(" ".join(chunk_words))
                file_idx += 1
                files_created += 1
        except Exception as e:
            messagebox.showerror("Error", f"Could not write split file:\n{e}")
            return


        messagebox.showinfo("Success", f"Text file split into {files_created} files in directory:\n{dir_name}")


    def replace_words_in_text_file(self):
        """Replaces words/phrases using a hybrid specific-then-general approach with refined context matching."""
        txt_path = self.text_file_path.get()
        if not os.path.isfile(txt_path):
            messagebox.showerror("Error", "Selected text file does not exist.")
            return

        try:
            # Try multiple common encodings for text files
            text = None
            encodings_to_try = ['utf-8', 'cp1252', 'latin-1']
            for encoding in encodings_to_try:
                try:
                    with open(txt_path, "r", encoding=encoding) as f:
                        text = f.read()
                    print(f"Successfully read file '{os.path.basename(txt_path)}' with encoding: {encoding}")
                    break # Exit loop if successful
                except UnicodeDecodeError:
                    print(f"Failed to read with encoding: {encoding}")
                    continue # Try next encoding
                except Exception as e:
                    # Handle other file reading errors
                    messagebox.showerror("Error", f"Error reading file {os.path.basename(txt_path)}:\n{e}")
                    return

            if text is None:
                messagebox.showerror("Error", f"Could not decode text file {os.path.basename(txt_path)} using common encodings.")
                return


            changes_made = False
            processed_text = text

            # --- Pre-compute mappings for the hybrid logic ---

            # 1. Create a map for EXACT original rules (lowercase original -> exact new)
            exact_replacements = {old.lower(): new for old, new in WORDS_TO_REPLACE}
            # print("DEBUG: Exact replacements map created.") # Optional debug

            # 2. Create a map for Base Phrases to their representative replacement word/phrase
            # This map defines the core transformation (e.g., "fuck" -> "Fetch", "piece of crap" -> "Piece of trash")
            # Derived from the first occurrence in the original list.
            base_transformations = {}
            for old, new in WORDS_TO_REPLACE:
                base_phrase = old.strip().lower() # Get the core offensive phrase (stripped)
                representative_replacement = new.strip() # Get the core replacement phrase (stripped of surrounding context *from the new string*)
                # Handle cases where the new string might be empty or just space, e.g., ("Shitless ", " ")
                if not representative_replacement:
                    representative_replacement = new # Use the original new string if stripped is empty

                if base_phrase not in base_transformations:
                    base_transformations[base_phrase] = representative_replacement
            # print("DEBUG: Base transformations map created.") # Optional debug


            # 3. Get unique base phrases and sort them by length descending for processing order
            sorted_base_phrases = sorted(base_transformations.keys(), key=len, reverse=True)
            # print(f"DEBUG: Processing order of base phrases: {[p for p in sorted_base_phrases]}") # Optional debug


            # --- Apply replacements using hybrid logic ---

            # Iterate through the sorted unique base phrases
            for base_phrase in sorted_base_phrases:
                # Compile a refined flexible regex pattern for this base phrase
                # It captures optional leading whitespace/quotes (Group 1)
                # Matches the escaped base phrase (Group 2) - using re.escape for safety
                # It captures optional trailing whitespace/non-word chars/quotes (Group 3)
                # This pattern is designed to capture the base phrase when it functions as a word unit
                pattern = re.compile(rf"([\s\"']*)({re.escape(base_phrase)})([\s\W\"']*)", re.IGNORECASE)

                # Define the custom replacer function here so it has access to pre-computed maps
                def hybrid_replacer(m):
                    original_matched_text = m.group(0) # The full text matched by the flexible pattern (e.g., " Fuck! ")
                    captured_leading_context = m.group(1) # Captured leading context (space/quotes)
                    matched_base_phrase_part = m.group(2) # The core base phrase matched by the pattern (e.g., "Fuck")
                    captured_trailing_context = m.group(3) # Captured trailing context (space/punct/quotes)

                    # --- HYBRID REPLACEMENT LOGIC ---
                    lowercase_matched_text = original_matched_text.lower()

                    if lowercase_matched_text in exact_replacements:
                        # Option 1: Found an EXACT match in the original list -> Use the exact replacement
                        final_replacement = exact_replacements[lowercase_matched_text]
                        # --- DEBUGGING PRINT ---
                        # print(f"DEBUG: Specific rule '{original_matched_text}' -> '{final_replacement}' matched exactly.")
                        # -----------------------
                    else:
                        # Option 2: No exact match -> Use general fallback
                        # Get the representative base replacement word/phrase for this base phrase
                        # Use the base_phrase from the *dictionary key* for consistent lookup casing
                        representative_replacement_word = base_transformations[base_phrase.lower()] # Use the key from the outer loop

                        # Reconstruct the replacement: original leading context + representative word + original trailing context
                        final_replacement = captured_leading_context + representative_replacement_word + captured_trailing_context
                         # --- DEBUGGING PRINT ---
                        # print(f"DEBUG: General fallback for '{original_matched_text}' (base '{matched_base_phrase_part}') using base '{representative_replacement_word}', result: '{final_replacement}'.")
                        # -----------------------
                    # --------------------------------

                    return final_replacement

                # Use the custom hybrid replacer function for substitution
                # We use subn to get the count for tracking changes and optional rule-specific logging
                # Use the replacer function with the compiled pattern
                new_text, count = pattern.subn(hybrid_replacer, processed_text)


                if count > 0:
                    changes_made = True # Mark that changes occurred
                    processed_text = new_text
                    # Optional: Print which base phrase rule caused *any* changes (Less verbose than inside replacer DEBUG)
                    # print(f"DEBUG: Base phrase rule for '{base_phrase}' caused {count} changes.")


            if changes_made:
                # Build output path: add _replaced before extension
                base, ext = os.path.splitext(txt_path)
                if not ext: ext = ".txt" # Default to .txt if no extension
                out_filename = f"{base}_replaced{ext}"
                out_path = os.path.join(os.path.dirname(txt_path), out_filename)


                # Simple overwrite prevention
                counter = 1
                original_out_name_for_msg = out_filename # Use this for user message if conflict occurs
                while os.path.exists(out_path):
                    out_filename = f"{base}_replaced_{counter}{ext}"
                    out_path = os.path.join(os.path.dirname(txt_path), out_filename) # Corrected path joining
                    counter += 1
                if counter > 1:
                    print(f"Warning: Output file '{original_out_name_for_msg}' existed, saving as '{os.path.basename(out_path)}'")


                with open(out_path, "w", encoding="utf-8") as f:
                    f.write(processed_text)
                messagebox.showinfo("Success", f"File saved with replacements as:\n{out_path}")
            else:
                messagebox.showinfo("Info", f"No words/phrases from the list were found in the text file. No new file saved.")

        except Exception as e:
            # This will now catch and display any other errors during the process
            messagebox.showerror("Error", f"Error processing text file:\n{e}")


    # --- Generate Samples Method ---
    def generate_samples(self):
        """Generates sample audio files for English voices with various speed/pause settings."""
        try:
            project_root = get_project_root() # Call the function
            # Use os.path.join for robustness
            input_path_sample = os.path.join(project_root, "input", "DO NOT DELETE", "DO NOT DELETE.txt") # Full path to sample text
            output_dir_samples = os.path.join(project_root, "output", "samples") # Output directory for samples
            os.makedirs(output_dir_samples, exist_ok=True) # Ensure output directory exists


            if not os.path.isfile(input_path_sample):
                messagebox.showerror("Error", f"Sample text file not found:\n{input_path_sample}")
                return

            # Check for piper executable
            # Use os.path.join for robustness
            piper_executable_abs = os.path.join(project_root, "piper_win", "piper.exe")
            if not os.path.isfile(piper_executable_abs):
                messagebox.showerror("Error", f"Piper executable not found:\n{piper_executable_abs}")
                return


            length_scales = [0.6, 0.7, 0.8, 0.9, 1.00]
            sentence_silences = [0.05, 0.1, 0.15, 0.2, 0.25, 0.30]

            print("Generating voice samples...")
            # --- Pre-process paths that are constant within the loop ---
            input_path_sample_bash = input_path_sample.replace("\\", "/")
            piper_executable_bash = piper_executable_abs.replace("\\", "/")
            # -----------------------------------------------------------


            for display_name, model_relpath in VOICES:
                # Only generate for English voices as previously decided/coded
                if not display_name.startswith("en_"):
                    continue
                model_path_abs = os.path.join(project_root, model_relpath) # Absolute path to model file
                if not os.path.isfile(model_path_abs):
                    print(f"Skipping sample for missing model file: {model_relpath}")
                    continue # Skip this voice if model is missing

                base_model = os.path.splitext(os.path.basename(model_path_abs))[0]

                # --- Pre-process model path for bash command (constant for ls/ss loop) ---
                model_path_bash = model_path_abs.replace("\\", "/")
                # ----------------------------------------------------------------------

                for ls in length_scales:
                    for ss in sentence_silences:
                        output_file_name = f"{base_model}_ls{ls:.2f}_ss{ss:.2f}.wav"
                        output_path_sample = os.path.join(output_dir_samples, output_file_name) # Full path to output wav

                        # --- Pre-process output path for bash command ---
                        output_path_sample_bash = output_path_sample.replace("\\", "/")
                        # --------------------------------------------------

                        # Construct the command using bash -c for robust path handling
                        # Outer bash command uses double quotes
                        # Inner double quotes around paths need to be escaped for bash (\")
                        # The backslash used for bash escaping needs to be escaped for the Python string (\\")
                        command = (
                            f'bash -c "MSYS2_ARG_CONV_EXCL=\\"*\\" cat \\"{input_path_sample_bash}\\" | ' # Use pre-processed path
                            f'\\"{piper_executable_bash}\\" ' # Use pre-processed path
                            f'--model \\"{model_path_bash}\\" --length_scale {ls:.2f} --sentence_silence {ss:.2f} ' # Use pre-processed path
                            f'--output_file \\"{output_path_sample_bash}\\""' # Use pre-processed path
                        )

                        # print(f"Executing: {command}") # Uncomment for debugging the command string
                        # Use shell=True is necessary with bash -c and subprocess.run on Windows
                        result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='utf-8')
                        # print("STDOUT:", result.stdout) # Uncomment for debugging subprocess output
                        # print("STDERR:", result.stderr) # Uncomment for debugging subprocess output

                        if result.returncode != 0:
                            print(f"ERROR generating {output_path_sample}:\nSTDOUT: {result.stdout}\nSTDERR: {result.stderr}")
                            # Optionally, show a non-blocking message or log this error
                        # else:
                        #   print(f"Generated: {output_path_sample}") # Too verbose


            messagebox.showinfo("Done", f"Voice samples generated in directory:\n{output_dir_samples}")
            print("Sample generation complete.")

        except Exception as e:
            messagebox.showerror("Error", f"Sample generation failed: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PiperTTSGUI(root)
    root.mainloop()