def leafNodeDescription(CMID,sellername,ticketType):

    if ticketType=="leafNode":
        issue="has absence of leaf node issue"

    elif ticketType=="keyword":
        issue="has no keywords"

    elif ticketType == "bulletpoints":
        issue = "has absence of 1-2-3 bullet point"

    elif ticketType == "title":
        issue = "has not product titles start with brand name"

    elif ticketType == "image":
        issue = "has absence of main image, due to these circumstances these SKU's suppressed"

    elif ticketType == "product_description":
        issue = "has no product description"

    leafNodeDescription=f"""CID:{CMID}
Seller Name: {sellername}
Issue desc:
Listed ASIN's in the attachments {issue}. Please take necessary actions.

Regards,
IROH The Ticket Operator

**Notes from a resolver**
For prioritization and escalation please reach out to @ebengisu
    
**The request will be handled within SLA based on severity as;**
Sev-5 Ticket SLA: 10 Business Days
Sev-4 Ticket SLA: 5 Business Days
Sev-3 Ticket SLA: 3 Business Days
    
Please visit below wiki page to view list of tasks covered under IDQ.
https://w.amazon.com/bin/view/TR_3P_IDQ 
    
**Tasks under the SIM requests by AMs:**
**!Tasks in the Issues Template area in the SIM folder which will be added to be selected by requesters!**

IDQ Related requests:

Keyword Backfilling 
Leaf Node Backfilling
Title Correction
Product Description Correction
Bullet Point Correction
GL Change
Detail Page/Attribute Correction
Category and Subcategory Backfilling
    """
    return leafNodeDescription