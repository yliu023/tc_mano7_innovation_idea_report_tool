def main()
    construct_datetimes_range()
    ideasSet=set()
    duplicatedList=list()
    notCalcIdeaList=list()
    #add alll files to data containers
    for file in files:
        addData(file)    
        #verify the data and construct CIdea instance    
        for i in range(len(idea_numbers)):
            idea=CIdea(idea_numbers[i],idea_titles[i], idea_initialtors[i],idea_funnels[i], idea_status[i], 
                   community_discussion[i],
                   implementation_started[i],
                   implementation_done[i]
                  )

            if idea.isInDateRange() and idea.isValid():
                if  ideasSet.add(idea):
                    duplicatedList.append(idea.getNumber())
            else:    
                notCalcIdeaList.append(idea.getNumber())                

    print("Total:", len(ideasSet), "Duplicated ideas:", len(duplicatedList), "Invalid ideas:", len(notCalcIdeaList))

    for idea in ideasSet:
        #print(idea.getNumber())
        SCORES[int(idea.statusIndex())][int(idea.teamIndex())] = SCORES[int(idea.statusIndex())][int(idea.teamIndex())] + 1

    # calc the total score for each team
    for i in range(len(SCORES[0])):
        SCORES[3][i]=SCORES[1][i]+SCORES[2][i]*3

    print(SCORES)
    saveStatisticCsv(statisticCsv, SCORES, ideasSet)
    testChart(compactCsv, SCORES)
