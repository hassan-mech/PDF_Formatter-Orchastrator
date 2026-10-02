# Let's inspect the exact lines from the images of pages 4 to 13
# We can do this by checking each page's blocks or by carefully reconstructing the text.

text_pages = {}

# Page 4
text_pages[4] = {
    'heading': 'PAUL STANLEY KC :',
    'items': [
        ('p', 1, 'The matter listed before me this morning is a summary judgment application. The claimant (which I shall call "Invest") seeks summary judgment under a guarantee given by the defendant (which I shall call "Lewa Trading"). Lewa Trading has not attended the application. It is not represented. It has not said anything to indicate its position formally or informally. I must decide whether to proceed with the hearing in the absence of the defendant. I am satisfied, as I told Mr. Oakes at the outset of the hearing, that I should do so.'),
        ('p', 2, 'I take into account in reaching that decision the factors identified in Butcher J in European Union v Syria [2023] EWHC 1116, and by Lionel Persey KC in African Export-Import Bank v National Government of the Republic of South Sudan [2025] EWHC 1079 (Comm). I am satisfied that the steps taken to bring this application to Lewa Trading\'s attention are likely to have done so. The application and the information about its listing have been notified to it both through the agent that it appointed to receive court process and through various e-mail addresses, and by post to its offices in Saudi Arabia, in sufficient time to make it likely that it will have been aware of the hearing. It was notified of this application on 27th May 2026 and of the date for the hearing on 2nd June 2026. Its lack of any response seems more likely to reflect a deliberate decision not to engage with proceedings than ignorance.'),
        ('p', 3, 'There is a strong public interest in defendants being notified of hearings in good time so they can participate in them, but there is equally a strong public interest in the orderly determination of claimants\' rights. The purpose of notifying defendants of hearings is to enable them to participate if they chose to do so. They are not compelled to do so, but their voluntary decision not to cannot stop the wheels of justice in their tracks; and is no less important where the amounts at stake are, as they are here, substantial. I am satisfied, therefore, there are compelling reasons for the hearing to proceed and no injustice to Lewa Trading if it does. Adjournment would serve no useful purpose.'),
        ('note', 0, '(For continuation of proceedings: please see separate transcript)'),
        ('p', 4, 'The application before me is for summary judgment. At the opening of the hearing today, I explained my reasons for proceeding with the hearing in the absence of the defendant, which I am satisfied has notice of it.'),
        ('p', 5, 'The case arises out of a Facility Agreement dated 2nd December 2022. The parties to that agreement were Naylor Nutrition UK Limited as borrower, Naylor Nutrition Group Limited as Guarantor, and the claimant (which I call "Invest"), which is a Dutch company, as lender.'),
        ('p', 6, 'In the evidence before me Naylor Nutrition Group Limited has been called "Naylor Guarantor" and I shall follow that practice too.'),
        ('p', 7, 'The agreement was a term Facility Agreement to provide export finance credit in the sum of €34.6 million.')
    ]
}

print("Page 4 prepared with", len(text_pages[4]['items']), "items")
