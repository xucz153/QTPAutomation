'Method info: Select  a item in Web Edit  randomly 
'Input  value
'PageObj:Object of the page
'ClassValue: class value of  children object
'GetChildValue: the value of displayed properties
'SearchMark is marked when in search test, the last  n links should be ingored
'SuggestMark is marked when the first  n links should be ingored
'Output  value
'Get a item in Web Edit  randomly 
'Created by ashley.shen, created time: 7/25/2012

'Get the displayed items in 1in Web Edit  list 
Public function GetItemOptionArr2(PageObj2, IValue, IValue1, GetChildValue2,SuggestMark,SearchMark)
	Set desc= description.Create
    desc("micclass").value=IValue
	desc("html tag").value=IValue1
	wait(environment.Value("sTime"))
	Set ProductItems= PageObj2.ChildObjects(desc)
	wait(environment.Value("sTime"))
    Redim OptionsArr(ProductItems.count)
	str=""
	'Only get the suggested items
		For i=SuggestMark to ProductItems.count-1-SearchMark
	
			OptionsArr(i)=ProductItems(i).GetRoProperty(GetChildValue2)
		Next
		GetItemOptionArr2=OptionsArr
	
End Function


'Method info: Select  a item in Web Edit  randomly 
'Input  value
'PageObj:Object of the page
'ClassValue: class value of  children object
'GetChildValue: the value of displayed properties
'SearchMark is marked whether the last  n links should be ingored
'SuggestMark is marked whether the first  n links should be ingored
'Output  value
'Created by ashley.shen, created time: 7/25/2012

'Select the 1st product item in 1st page of  Batteries product list  randomly
Public function GetRndSuggestedtem(PageObj2, IValue, IValue1, GetChildValue2,SuggestMark,SearchMark)
	OptionsArr2=GetItemOptionArr2(PageObj2, IValue, IValue1, GetChildValue2,SuggestMark,SearchMark)
    Randomize
	c= int((ubound(OptionsArr2)-SuggestMark-SearchMark)*rnd+SuggestMark)
	print "c and OptionsArr2(c) is" &c &","& OptionsArr2(c)
    GetRndSuggestedtem=OptionsArr2(c)
END FUNCTION

RegisterUserFunc "WebElement","GetItemOptionArr2","GetItemOptionArr2",True
RegisterUserFunc "WebElement","GetRndSuggestedtem","GetRndSuggestedtem",True
