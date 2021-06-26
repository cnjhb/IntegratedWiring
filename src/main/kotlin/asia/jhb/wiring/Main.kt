package asia.jhb.wiring
import com.alibaba.excel.EasyExcel
import com.alibaba.excel.annotation.ExcelProperty
import com.alibaba.excel.context.AnalysisContext
import com.alibaba.excel.event.AnalysisEventListener
import java.io.File

private val DISTANCEPOINTS = 10
private val DISTANCEFLOORS = 4
fun main(args:Array<String>) {
    val target = File(args[0])
    if(!target.exists()||!target.isFile) {
        println("文件${target.absoluteFile}不存在")
        help()
        return
    }
    var result = "result.xlsx"
    for(index in args.indices){
        if(args[index]=="-o")
        {
            try {
                result = args[index + 1]
            }catch (e:ArrayIndexOutOfBoundsException){
                help()
            }
        }
    }
    var template = Layout().javaClass.getResourceAsStream("/template.xlsx")
    var wiring = 0
    var points = 0
    var floors = 0
    EasyExcel.read(target, Layout::class.java, object : AnalysisEventListener<Layout>() {
        override fun invoke(p0: Layout, p1: AnalysisContext) {
            points += p0.points
            wiring += pointToLength(p0.points)
            floors++
        }

        override fun doAfterAllAnalysed(p0: AnalysisContext?) {
            wiring+= floorToLength(floors)
        }

    }).sheet().doRead()
    val fillLayout = FillLayout(points, wiring,floors)
    EasyExcel.write(result).withTemplate(template).sheet().doFill(fillLayout)
}

data class Layout(@ExcelProperty("信息点数") var points: Int = 0, @ExcelProperty("层数") var floorName: String = "")

fun pointToLength(points: Int): Int {
    var length = 0
    for (i in 0 until points / 2) {
        length += i * DISTANCEPOINTS
    }
    return length * 2
}

fun floorToLength(floors: Int): Int {
    var length = 0
    for (i in 0..floors)
        length += i * DISTANCEFLOORS
    return length
}

fun help(){
    println("""
        使用方式:程序 目标文件 【参数】
        参数：
        -o 输出到的文件
    """.trimIndent())
}

data class FillLayout(var points: Int, var wiring: Int,var floors: Int)