module CUSTODESGrammar
    open FParsec
    open System.Collections.Generic

    let private DEBUG_MODE = false

    // a convenient type alias
    type P<'t> = Parser<'t, unit>  

    // a debug parser
    let private (<!>) (p: P<_>) label : P<_> =
        #if DEBUG
            fun stream ->
                let s = String.replicate (int (stream.Position.Column)) " "
                let pre = sprintf "%d%sTrying %s(\"%s\")" (stream.Index) s label (stream.PeekString 1000)
                System.Diagnostics.Debug.WriteLine(pre)
                let reply = p stream
                let post = sprintf "%d%s%s(\"%s\") (%A)" (stream.Index) s label (stream.PeekString 1000) reply.Status
                System.Diagnostics.Debug.WriteLine(post)
                reply
        #else
            p 
        #endif

    // CUSTODES output datatypes
    type Worksheet = string
    type Address = string
    type CUSTODESSmells = Dictionary<Worksheet,Address[]>

    let private analysisStart : P<Worksheet> =
        between
            (pstring "----procesing worksheet '")
            (pstring "'----" .>> spaces)
            (manyChars (noneOf "'"))
        <!> "analysisStart"

    let private clusterStart : P<string> =
        pstring "---- Stage I clustering begined ----" .>> spaces
        <!> "clusterStart"
    let private clusterEnd : P<string> =
        pstring "---- Stage I finished ----" .>> spaces
        <!> "clusterEnd"
    let private clustersFound : P<int> =
        between
            (pstring "found ")
            (pstring " clusters" .>> spaces)
            pint32
        <!> "clustersFound"
    let private clusters : P<int> =
        clusterStart >>. (clusterEnd >>. clustersFound)
        <!> "clusters"

    let private smellStart : P<string> =
        pstring "---- Stage II begined ----" .>> spaces
        <!> "smellStart"
    let private smellsFound : P<int> =
        between
            (pstring "detected ")
            (pstring " smells:" .>> spaces)
            pint32
        <!> "smellsFound"
    let private smellCell : P<Address> =
        many1Satisfy (fun c -> isLetter(c) || isDigit(c)) .>> spaces
        <!> "smellCell"
    let private smellCells : P<Address[]> =
        many smellCell |>> (fun cells -> cells |> List.toArray)
        <!> "smellCells"
    let private smellEnd : P<string> =
        pstring "---Analysis Finished---" .>> spaces
        <!> "smellEnd"
    let private smells : P<Address[]> =
        smellStart >>. ((attempt smellsFound) >>. (smellCells .>> smellEnd))
        <!> "smells"

    let private worksheetSomeSmells : P<CUSTODESSmells> =
        pipe2
            (analysisStart .>> clusters)
            smells
            (fun w smells ->
                let d = new CUSTODESSmells()
                d.Add(w, smells)
                d
            )
        <!> "worksheetSomeSmells"

    let private worksheetNoSmells : P<CUSTODESSmells> =
        analysisStart
            |>>
            (fun w ->
                let d = new CUSTODESSmells()
                d.Add(w, [||])
                d
            )
        <!> "worksheetNoSmells"

    let private worksheet : P<CUSTODESSmells> =
        (attempt worksheetSomeSmells) <|> worksheetNoSmells
        <!> "worksheet"

    let private worksheets : P<CUSTODESSmells> =
        (many1 worksheet |>>
            (fun (cslist: CUSTODESSmells list) ->
                let d = new CUSTODESSmells()
                for cs in cslist do
                    for pair in cs do
                        d.Add(pair.Key, pair.Value)
                d
            )
        )
        <!> "worksheets"

    let private start : P<CUSTODESSmells> =
        worksheets .>> eof
        <!> "start"

    let exceptionParser : P<string> =
        skipManyTill anyChar (pstring "Exception in thread") >>.
            (
                (many1 anyChar) |>>
                (fun stacktrace -> "Exception in thread " + System.String.Join("", stacktrace))
            )
        <!> "exceptionParser"

    type CUSTODESParse =
    | CSuccess of CUSTODESSmells
    | CFailure of string

    let parseException(output: string) : CUSTODESParse =
        match run exceptionParser output with
        | Success(excptn,_,_) -> CFailure(excptn)
        | Failure(other_failure,_,_) -> CFailure(other_failure)

    let parse(output: string) : CUSTODESParse =
        // before parsing, look for exceptions
        match run exceptionParser output with
        | Success(excptn,_,_) -> CFailure(excptn)
        | Failure(noexcptn,_,_) -> 
            // good, now parse output
            match run start output with
            | Success(smells,_,_) -> CSuccess(smells)
            | Failure(err,_,_) -> CFailure(err)