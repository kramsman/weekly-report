# test parsing state/county from campian name

parent_campaign_name = "VA-Arlington-CD8 Primary GOTV 4-2022 **ASAP by JUNE 6**"
s = parent_campaign_name.split("-")
statecnty = parent_campaign_name.split("-")[0]+"-"+parent_campaign_name.split("-")[1]
a=1
