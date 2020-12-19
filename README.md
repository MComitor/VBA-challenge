### VBA-challenge ###

Lots of challenges with this exercise! 

<details>
    <summary>Easy Solution</summary>
        <p>First of all, I experienced issues with my "easy" solution where the output was **NOT** color formatted. I struggled to update the code multiple times by commenting out the For loop section in lines 65 and 76 of the easy solution, and just leaving behind the If ElseIf statement. Yet this caused the entire Yearly Change (Column J) to fill with Green. However, upon re-running the code again a second time without clearing the output, then the color formatting is applied appropriately throughout column J. But this method is not quite right since the code should only run once to yield the expected resutls. Furthermore, when I attempted to leave the For Loop intact, no color formatting is applied and I get an 'Overflow' error in line 61 (percent_change), likely due to a Divide by Zero error..? But why didn't this error occur when I ran it the first time?! I'm not sure where I went wrong here. I've included two images to capture the behavior I experienced: One image of the first run which produced no color formatting when I left the For loop intact (easy_solution_image_run1) and then a second run through of the script and the color formatting applied appropriately across Column J (easy_solution_image_run2) when the script was re-run a second time.</p>
</details>

<details>
    <summary>Moderate Solution</summary>
        <p>Next up, I experienced issues with my "moderate" solution. While this script did generate a new Combined Sheet, all sheet values failed to carry over. It's like the script gets hung up on the first ticker line and then doesn't know how to proceed. The debugger highlighted line 62 as a 'Type mismatch' error (stock_vol). This is confusing to me because I've defined the stock_volume as LongLong initally. What's happening here?</p>
</details>

<details>
    <summary>Bonus</summary>
        <p>I ended up abandoning this extra credit portion of the assignment as I couldn't even get my easy solution to work successfully.</p>
</details> 