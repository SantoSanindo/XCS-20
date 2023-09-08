Module Declaration
    Public LoadWOfrRFID As JobOrder
    Public rprintlabel As JobOrder
    Public LabelVar As LabelData 'Used in the main page
    Public ReprintlabelVar As LabelData 'Used in the reprint page
    Public DUTParameter As ProductData
    Public Unitmaterial As RackConfig
    Public PSNFileInfo As PSNText
    Public Parameter As ControlSpec
    Public TestRead As ReadbackContact
    Public Structure JobOrder
        Dim JobNos As String
        Dim JobModelName As String
        Dim JobQTy As Integer
        Dim JobArticle As String
        Dim JobModelFW As String
        Dim JobModelAssy As String
        Dim JobInternalNos As String
        Dim JobRFIDTag As String
        Dim JobUnitaryCount As Integer
        Dim JobModelMaterial() As String
        Dim JobProductMaterial As String 'Zamak or Plastic
        Dim JobProductThread As String
        Dim JobButtonType As String 'With PushButton?
        Dim JobConnectorType As String
        Dim JobConnectorID As String
    End Structure
    Public Structure LabelData
        Dim UnitModelName As String
        Dim UnitRefLogistique As String
        Dim UnitArticleNos As String
        Dim UnitDetail1 As String 'String
        Dim UnitDetail2 As String
        Dim UnitDetail3 As String
        Dim UnitDetail4 As String
        Dim UnitDetail5 As String
        Dim UnitDetail6 As String
        Dim UnitDetail7 As String
        Dim UnitDetail8 As String
        Dim UnitDetail9 As String
        Dim UnitDetail10 As String
        Dim UnitImage As String
        Dim UnitLabelTemplate As String
        Dim UnitLabelSymb1 As String
        Dim UnitLabelSymb2 As String
        Dim UnitLabelTor As String
    End Structure
    Public Structure ProductData
        Dim UnitTension As String
        Dim UnitFunctionType As String
        Dim UnitS11ContactNos As String
        Dim UnitS12ContactNos As String
        Dim UnitS21ContactNos As String
        Dim UnitS22ContactNos As String
        Dim UnitS31ContactNos As String
        Dim UnitS32ContactNos As String
        Dim UnitS41ContactNos As String
        Dim UnitS42ContactNos As String
        Dim UnitS51ContactNos As String
        Dim UnitS52ContactNos As String
        Dim UnitS61ContactNos As String
        Dim UnitS62ContactNos As String
        Dim UnitS11PN As String
        Dim UnitS12PN As String
        Dim UnitS21PN As String
        Dim UnitS22PN As String
        Dim UnitS31PN As String
        Dim UnitS32PN As String
        Dim UnitS41PN As String
        Dim UnitS42PN As String
        Dim UnitS51PN As String
        Dim UnitS52PN As String
        Dim UnitS61PN As String
        Dim UnitS62PN As String
        Dim UnitX1PN As String
        Dim UnitX2PN As String
        Dim UnitE1PN As String
        Dim UnitE2PN As String
        Dim UnitS1CC As String
        Dim UnitS2CC As String
        Dim UnitS3CC As String
        Dim UnitS4CC As String
        Dim UnitS5CC As String
        Dim UnitS6CC As String
        Dim UnitX1CC As String
        Dim UnitX2CC As String
        Dim UnitE1CC As String
        Dim UnitE2CC As String
        Dim UnitContact_original() As String
        Dim UnitContactOriginaltemp() As String
        Dim UnitContact_Wkey() As String
        Dim UnitcontactWkeytemp() As String
        Dim UnitContact_WkeyTension() As String
        Dim UnitContactWkeyTensiontemp() As String
        Dim UnitOriginalSate As String
        Dim UnitStatewkey As String
        Dim UnitStatewkeyntension As String
    End Structure
    Public Structure RackConfig
        Dim PartPLCWord() As Long
        Dim PartNos() As String
    End Structure
    Public Structure PSNText
        Dim MODELNAME As String
        Dim DateCreated As String
        Dim DateCompleted As String
        Dim OperatorID As String
        Dim WONos As String
        Dim MainPCBA As String
        Dim SecondaryPCBA As String
        Dim ElectroMagnet As String
        Dim PSN As String
        Dim BodyAssyCheckIn As String
        Dim BodyAssyCheckOut As String
        Dim BodyAssyStatus As String
        Dim ScrewStnCheckIn As String
        Dim ScrewStnCheckOut As String
        Dim ScrewStnStatus As String
        Dim FTCheckIn As String
        Dim FTCheckOut As String
        Dim FTSycMeas As String
        Dim FTStatus As String
        Dim Stn5CheckIn As String
        Dim Stn5CheckOut As String
        Dim Stn5Status As String
        Dim VacuumCheckIn As String
        Dim VacummCheckOut As String
        Dim VacuumStatus As String
        Dim ConnTestCheckIn As String
        Dim ConnTestCheckOut As String
        Dim ConnTestStatus As String
        Dim Vacuum2CheckIn As String
        Dim Vacumm2CheckOut As String
        Dim Vacuum2Status As String
        Dim PackagingCheckIn As String
        Dim PackagingCheckOut As String
        Dim PackagingStatus As String
        Dim DebugStatus As String
        Dim DebugComment As String
        Dim DebugTechnican As String
        Dim RepairDate As String
    End Structure
    Public Structure ControlSpec
        Dim UnitTagNos As String
        Dim UnitModel As String
    End Structure
    Public Structure ReadBackContact
        Dim OriginState() As String
        Dim KeyState() As String
        Dim KeyTension() As String
    End Structure
End Module
