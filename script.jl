using RCall
using CSV

const INPFOLDER = "CSV"
const OUTFOLDER = "RDATA"

rfy(x) = string("c('", join(x, "','"), "')")

function export_rda(dir)
    """
    Export CSV output from Python to R's rda
    """
    csv_files = readdir(dir)
    
    for csv in csv_files
        data = CSV.read(joinpath(INPFOLDER, csv))
        code = map(x -> x === missing ? "NA" : x, data[:Code]) |> rfy
        s1   = map(x -> x === missing ? "NA" : x, data[:S1]) |> rfy
        s2   = map(x -> x === missing ? "NA" : x, data[:S2]) |> rfy
        s3   = map(x -> x === missing ? "NA" : x, data[:S3]) |> rfy
        s11  = map(x -> x === missing ? "NA" : x, data[:S1_1]) |> rfy
        s22  = map(x -> x === missing ? "NA" : x, data[:S2_1]) |> rfy
        s33  = map(x -> x === missing ? "NA" : x, data[:S3_1]) |> rfy
        wts  = map(x -> x === missing ? "NA" : x, data[Symbol("Weight class")]) |> rfy
        
        crop_name = split(csv, ".")[1]
        outfile = joinpath(OUTFOLDER, crop_name * ".rda")
        rcode = """
            $crop_name <- data.frame(code = $code, s3_a = $s3, s2_a = $s2, s1_a = $s1, s1_b = $s11, s2_b = $s22, s3_b = $s33, wts = $wts)
            $crop_name <- replace($crop_name, $crop_name == 'NA', NA)
            save($crop_name, file='$outfile')
        """
        reval(rcode)
    end
end


export_rda("CSV")