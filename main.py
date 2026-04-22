from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st


DATA_PATH = Path(__file__).parent / "data" / "masters_admission_sample.csv"

GRE_POLICY_LABELS = {"No": "不要求", "Optional": "可选", "Required": "必需"}
MAJOR_REQUIREMENT_LABELS = {
    "Open": "接受跨专业",
    "Medium": "偏好相关背景",
    "Strong": "强相关背景",
}
TIER_BONUS = {
    "普通本科": 0,
    "双一流/211": 2,
    "985": 3,
    "C9/海外名校": 4,
}
MAJOR_RELEVANCE_SCORE = {
    "完全相关": {"Open": 100, "Medium": 100, "Strong": 100},
    "相近专业": {"Open": 90, "Medium": 82, "Strong": 68},
    "跨专业": {"Open": 78, "Medium": 55, "Strong": 32},
}
RECOMMENDATION_ORDER = {"保底": 0, "匹配": 1, "冲刺": 2, "暂不达标": 3}
RECOMMENDATION_COLORS = {
    "保底": "#2ca02c",
    "匹配": "#1f77b4",
    "冲刺": "#ff7f0e",
    "暂不达标": "#d62728",
}


st.set_page_config(
    page_title="海外硕士录取标准分析模型",
    page_icon="🎓",
    layout="wide",
)


@st.cache_data
def load_program_data() -> pd.DataFrame:
    program_data = pd.read_csv(DATA_PATH)
    numeric_columns = [
        "qs_rank",
        "tuition_usd",
        "duration_months",
        "min_gpa_100",
        "min_ielts",
        "min_toefl",
        "min_gre",
        "quant_requirement",
        "programming_requirement",
        "internship_importance",
        "research_importance",
    ]
    for column in numeric_columns:
        program_data[column] = pd.to_numeric(program_data[column], errors="coerce")

    program_data = program_data.dropna().copy()
    program_data["display_name"] = program_data["university"] + "｜" + program_data["program_name"]
    program_data["gre_policy_cn"] = program_data["gre_policy"].map(GRE_POLICY_LABELS)
    program_data["major_requirement_cn"] = program_data["major_requirement"].map(
        MAJOR_REQUIREMENT_LABELS
    )
    return program_data


def clip_score(value: float) -> float:
    return max(0.0, min(100.0, value))


def competition_penalty(qs_rank: float) -> int:
    if qs_rank <= 10:
        return 8
    if qs_rank <= 30:
        return 6
    if qs_rank <= 60:
        return 4
    if qs_rank <= 100:
        return 2
    return 0


def level_score(user_level: int, required_level: int) -> float:
    difference = user_level - required_level
    if difference >= 1:
        return 100
    if difference == 0:
        return 82
    if difference == -1:
        return 60
    return 35


def language_score(profile: dict, row: pd.Series) -> float:
    ielts_score = 55 + (profile["ielts"] - row["min_ielts"]) * 18
    toefl_score = 55 + (profile["toefl"] - row["min_toefl"]) * 1.8
    return clip_score(max(ielts_score, toefl_score))


def gre_score(profile: dict, row: pd.Series) -> float:
    if row["gre_policy"] == "No":
        return 80
    if profile["gre"] <= 0:
        return 58 if row["gre_policy"] == "Optional" else 0
    base = 60 if row["gre_policy"] == "Required" else 62
    return clip_score(base + (profile["gre"] - row["min_gre"]) * 2.2)


def calculate_components(row: pd.Series, profile: dict) -> dict:
    adjusted_gpa = profile["gpa"] + TIER_BONUS[profile["undergraduate_tier"]]
    gpa_component = clip_score(55 + (adjusted_gpa - row["min_gpa_100"]) * 5.5)
    language_component = language_score(profile, row)
    gre_component = gre_score(profile, row)
    hard_threshold_score = (
        gpa_component * 0.45 + language_component * 0.35 + gre_component * 0.20
    )

    major_component = MAJOR_RELEVANCE_SCORE[profile["major_relevance"]][
        row["major_requirement"]
    ]
    quant_component = level_score(profile["quant_level"], int(row["quant_requirement"]))
    programming_component = level_score(
        profile["programming_level"], int(row["programming_requirement"])
    )
    major_fit_score = (
        major_component * 0.45 + quant_component * 0.30 + programming_component * 0.25
    )

    internship_target = max(row["internship_importance"] * 2, 1)
    internship_component = min(profile["internship_months"] / internship_target, 1) * 100
    research_component = min(profile["research_count"] / max(row["research_importance"], 1), 1) * 100
    project_component = min(profile["project_count"] / 2, 1) * 100
    soft_power_score = (
        internship_component * 0.40 + research_component * 0.30 + project_component * 0.30
    )

    if profile["budget_usd"] >= row["tuition_usd"]:
        budget_score = 100
    elif profile["budget_usd"] >= row["tuition_usd"] * 0.85:
        budget_score = 75
    elif profile["budget_usd"] >= row["tuition_usd"] * 0.70:
        budget_score = 55
    else:
        budget_score = 30

    country_bonus = 3 if row["country"] in profile["preferred_countries"] else 0
    fit_score = (
        hard_threshold_score * 0.50
        + major_fit_score * 0.30
        + soft_power_score * 0.20
        + budget_score * 0.05
        + country_bonus
        - competition_penalty(row["qs_rank"])
    )

    return {
        "adjusted_gpa": round(adjusted_gpa, 1),
        "gpa_component": round(gpa_component, 1),
        "language_component": round(language_component, 1),
        "gre_component": round(gre_component, 1),
        "hard_threshold_score": round(hard_threshold_score, 1),
        "major_component": round(major_component, 1),
        "quant_component": round(quant_component, 1),
        "programming_component": round(programming_component, 1),
        "major_fit_score": round(major_fit_score, 1),
        "soft_power_score": round(soft_power_score, 1),
        "budget_score": round(budget_score, 1),
        "fit_score": round(clip_score(fit_score), 1),
    }


def build_reason(row: pd.Series, profile: dict, components: dict, hard_pass: bool) -> str:
    reasons = []
    if components["adjusted_gpa"] < row["min_gpa_100"]:
        reasons.append("GPA/院校背景未达到项目门槛")
    if profile["ielts"] < row["min_ielts"] and profile["toefl"] < row["min_toefl"]:
        reasons.append("语言成绩未达到最低要求")
    if row["gre_policy"] == "Required" and profile["gre"] < row["min_gre"]:
        reasons.append("GRE 为必需且当前未达标")
    if row["gre_policy"] == "Optional" and profile["gre"] <= 0:
        reasons.append("提交 GRE 可增强竞争力")
    if profile["major_relevance"] == "跨专业" and row["major_requirement"] != "Open":
        reasons.append("专业背景匹配度偏弱")
    if profile["quant_level"] < row["quant_requirement"]:
        reasons.append("数学/统计基础需要补强")
    if profile["programming_level"] < row["programming_requirement"]:
        reasons.append("编程基础需要补强")
    if profile["budget_usd"] < row["tuition_usd"]:
        reasons.append("项目学费高于当前预算")
    if hard_pass and not reasons:
        reasons.append("硬门槛满足，背景整体较均衡")
    return "；".join(reasons)


def calculate_match_results(program_data: pd.DataFrame, profile: dict) -> pd.DataFrame:
    result_rows = []
    for _, row in program_data.iterrows():
        components = calculate_components(row, profile)
        gpa_pass = components["adjusted_gpa"] >= row["min_gpa_100"]
        language_pass = (
            profile["ielts"] >= row["min_ielts"] or profile["toefl"] >= row["min_toefl"]
        )
        gre_pass = row["gre_policy"] != "Required" or profile["gre"] >= row["min_gre"]
        hard_pass = gpa_pass and language_pass and gre_pass

        if not hard_pass:
            recommendation = "暂不达标"
        elif components["fit_score"] >= 84:
            recommendation = "保底"
        elif components["fit_score"] >= 70:
            recommendation = "匹配"
        else:
            recommendation = "冲刺"

        result_rows.append(
            {
                **row.to_dict(),
                **components,
                "gpa_pass": gpa_pass,
                "language_pass": language_pass,
                "gre_pass": gre_pass,
                "hard_pass": hard_pass,
                "recommendation": recommendation,
                "reason": build_reason(row, profile, components, hard_pass),
            }
        )

    results = pd.DataFrame(result_rows)
    results["sort_key"] = results["recommendation"].map(RECOMMENDATION_ORDER)
    return results.sort_values(
        ["sort_key", "fit_score", "qs_rank"], ascending=[True, False, True]
    ).copy()


def build_sidebar(program_data: pd.DataFrame) -> tuple[dict, pd.DataFrame]:
    st.sidebar.header("申请人画像")
    undergraduate_tier = st.sidebar.selectbox(
        "本科背景",
        list(TIER_BONUS.keys()),
        index=1,
        help="用于模拟 985/211 等背景在部分学校中的门槛折算优势。",
    )
    profile = {
        "undergraduate_tier": undergraduate_tier,
        "gpa": st.sidebar.slider("本科 GPA（百分制）", 70, 95, 84),
        "major_relevance": st.sidebar.selectbox(
            "本科专业匹配度", ["完全相关", "相近专业", "跨专业"], index=1
        ),
        "quant_level": st.sidebar.slider("数学/统计基础", 1, 5, 3),
        "programming_level": st.sidebar.slider("编程基础", 1, 5, 3),
        "ielts": st.sidebar.slider("雅思", 0.0, 8.5, 6.5, 0.5),
        "toefl": st.sidebar.slider("托福", 0, 120, 0, 1),
        "gre": st.sidebar.number_input("GRE（未提交填 0）", 0, 340, 0, 1),
        "internship_months": st.sidebar.slider("相关实习时长（月）", 0, 24, 4),
        "research_count": st.sidebar.slider("科研/论文经历数量", 0, 5, 0),
        "project_count": st.sidebar.slider("数据/商业分析项目数", 0, 6, 1),
        "budget_usd": st.sidebar.slider("总学费预算（美元）", 20000, 75000, 45000, 1000),
        "preferred_countries": st.sidebar.multiselect(
            "目标国家/地区",
            sorted(program_data["country"].unique()),
            default=sorted(program_data["country"].unique()),
        ),
        "preferred_fields": st.sidebar.multiselect(
            "目标方向",
            sorted(program_data["field"].unique()),
            default=sorted(program_data["field"].unique()),
        ),
    }

    filtered_programs = program_data[
        program_data["country"].isin(profile["preferred_countries"])
        & program_data["field"].isin(profile["preferred_fields"])
    ].copy()
    return profile, filtered_programs


def show_metrics(results: pd.DataFrame, profile: dict) -> None:
    hard_pass_count = int(results["hard_pass"].sum())
    recommended_count = int(results["recommendation"].isin(["保底", "匹配"]).sum())
    budget_count = int((results["tuition_usd"] <= profile["budget_usd"]).sum())
    average_score = results["fit_score"].mean()
    columns = st.columns(5)
    columns[0].metric("筛选项目数", len(results))
    columns[1].metric("满足硬门槛", hard_pass_count)
    columns[2].metric("推荐申请", recommended_count)
    columns[3].metric("预算内项目", budget_count)
    columns[4].metric("平均匹配指数", f"{average_score:.1f}")


def show_auto_advice(results: pd.DataFrame) -> None:
    valid_results = results[results["recommendation"].isin(["保底", "匹配", "冲刺"])]
    if valid_results.empty:
        st.warning("当前申请画像下暂无可推荐项目，建议优先提升 GPA、语言成绩或放宽目标范围。")
        return

    top_programs = valid_results.head(3)
    best_program = top_programs.iloc[0]
    weak_dimensions = []
    component_means = {
        "硬门槛": best_program["hard_threshold_score"],
        "专业匹配": best_program["major_fit_score"],
        "软实力": best_program["soft_power_score"],
        "预算": best_program["budget_score"],
    }
    for name, value in component_means.items():
        if value < 65:
            weak_dimensions.append(name)
    weak_text = "、".join(weak_dimensions) if weak_dimensions else "暂无明显短板"

    st.info(
        f"当前最推荐关注 **{best_program['display_name']}**，结果为 **{best_program['recommendation']}**，"
        f"匹配指数 **{best_program['fit_score']}**。前三个可考虑项目为："
        f"{'；'.join(top_programs['display_name'].tolist())}。"
        f"主要需要关注的短板：**{weak_text}**。"
    )


def show_standards_tab(program_data: pd.DataFrame) -> None:
    st.subheader("项目录取标准总览")
    scatter = (
        alt.Chart(program_data)
        .mark_circle(size=170)
        .encode(
            x=alt.X("qs_rank:Q", title="QS 排名（数值越小竞争越强）"),
            y=alt.Y("min_gpa_100:Q", title="最低 GPA 门槛"),
            color=alt.Color("country:N", title="国家/地区"),
            size=alt.Size("tuition_usd:Q", title="学费"),
            tooltip=[
                alt.Tooltip("university:N", title="学校"),
                alt.Tooltip("program_name:N", title="项目"),
                alt.Tooltip("field:N", title="方向"),
                alt.Tooltip("min_gpa_100:Q", title="最低 GPA"),
                alt.Tooltip("min_ielts:Q", title="最低雅思"),
                alt.Tooltip("gre_policy_cn:N", title="GRE"),
                alt.Tooltip("tuition_usd:Q", title="学费", format=",.0f"),
            ],
        )
        .properties(height=380)
    )
    st.altair_chart(scatter, use_container_width=True)

    summary = (
        program_data.groupby("country", as_index=False)
        .agg(
            项目数=("program_name", "count"),
            平均最低GPA=("min_gpa_100", "mean"),
            平均最低雅思=("min_ielts", "mean"),
            平均学费USD=("tuition_usd", "mean"),
            平均QS排名=("qs_rank", "mean"),
        )
        .sort_values("平均QS排名")
        .round(1)
    )
    st.dataframe(summary, use_container_width=True, hide_index=True)


def show_match_tab(results: pd.DataFrame) -> None:
    st.subheader("个人匹配结果")
    distribution = (
        results["recommendation"]
        .value_counts()
        .rename_axis("recommendation")
        .reset_index(name="count")
    )
    pie_chart = (
        alt.Chart(distribution)
        .mark_arc(innerRadius=45)
        .encode(
            theta=alt.Theta("count:Q", title="数量"),
            color=alt.Color(
                "recommendation:N",
                scale=alt.Scale(
                    domain=list(RECOMMENDATION_COLORS.keys()),
                    range=list(RECOMMENDATION_COLORS.values()),
                ),
                title="结果分层",
            ),
            tooltip=["recommendation:N", "count:Q"],
        )
        .properties(height=300)
    )
    top_results = results.head(10)
    bar_chart = (
        alt.Chart(top_results)
        .mark_bar()
        .encode(
            x=alt.X("fit_score:Q", title="录取匹配指数"),
            y=alt.Y("display_name:N", sort="-x", title="项目"),
            color=alt.Color(
                "recommendation:N",
                scale=alt.Scale(
                    domain=list(RECOMMENDATION_COLORS.keys()),
                    range=list(RECOMMENDATION_COLORS.values()),
                ),
                title="建议",
            ),
            tooltip=[
                alt.Tooltip("display_name:N", title="项目"),
                alt.Tooltip("fit_score:Q", title="匹配指数"),
                alt.Tooltip("recommendation:N", title="建议"),
                alt.Tooltip("reason:N", title="原因"),
            ],
        )
        .properties(height=390)
    )
    left_column, right_column = st.columns([1, 2])
    with left_column:
        st.altair_chart(pie_chart, use_container_width=True)
    with right_column:
        st.altair_chart(bar_chart, use_container_width=True)

    display_columns = [
        "university",
        "country",
        "program_name",
        "field",
        "qs_rank",
        "tuition_usd",
        "min_gpa_100",
        "min_ielts",
        "gre_policy_cn",
        "major_requirement_cn",
        "fit_score",
        "recommendation",
        "reason",
    ]
    table = results[display_columns].rename(
        columns={
            "university": "学校",
            "country": "国家/地区",
            "program_name": "项目",
            "field": "方向",
            "qs_rank": "QS排名",
            "tuition_usd": "学费USD",
            "min_gpa_100": "最低GPA",
            "min_ielts": "最低雅思",
            "gre_policy_cn": "GRE要求",
            "major_requirement_cn": "专业要求",
            "fit_score": "匹配指数",
            "recommendation": "建议",
            "reason": "解释",
        }
    )
    st.dataframe(table, use_container_width=True, hide_index=True)


def show_model_tab(results: pd.DataFrame) -> None:
    st.subheader("模型解释：硬门槛 + 专业匹配 + 软实力")
    st.write(
        "本工具使用规则加权模型，先判断 GPA、语言、GRE 等硬门槛，再综合专业背景、数学/编程基础和软实力。"
        "结果用于课堂展示和初步选校，不代表学校官方录取概率。"
    )
    st.code("录取匹配指数 = 50%×硬门槛 + 30%×专业匹配 + 20%×软实力 + 预算修正 - 竞争度惩罚")

    selected_name = st.selectbox("选择项目查看得分拆解", results["display_name"].tolist())
    selected = results.loc[results["display_name"] == selected_name].iloc[0]
    component_frame = pd.DataFrame(
        {
            "维度": ["硬门槛", "专业匹配", "软实力", "预算"],
            "得分": [
                selected["hard_threshold_score"],
                selected["major_fit_score"],
                selected["soft_power_score"],
                selected["budget_score"],
            ],
        }
    )
    chart = (
        alt.Chart(component_frame)
        .mark_bar()
        .encode(
            x=alt.X("得分:Q", title="得分"),
            y=alt.Y("维度:N", sort="-x", title="维度"),
            color=alt.Color("维度:N", legend=None),
            tooltip=["维度:N", "得分:Q"],
        )
        .properties(height=300)
    )
    st.altair_chart(chart, use_container_width=True)

    columns = st.columns(4)
    columns[0].metric("匹配指数", selected["fit_score"])
    columns[1].metric("结果分层", selected["recommendation"])
    columns[2].metric("折算 GPA", selected["adjusted_gpa"])
    columns[3].metric("GRE 要求", selected["gre_policy_cn"])
    st.write(f"**解释：** {selected['reason']}")
    st.write(
        f"**项目门槛：** GPA {selected['min_gpa_100']}，雅思 {selected['min_ielts']}，"
        f"托福 {selected['min_toefl']}，专业背景：{selected['major_requirement_cn']}。"
    )


def show_cost_tab(program_data: pd.DataFrame, results: pd.DataFrame, profile: dict) -> None:
    st.subheader("费用与国家/地区对比")
    cost_chart = (
        alt.Chart(program_data)
        .mark_bar()
        .encode(
            x=alt.X("country:N", title="国家/地区"),
            y=alt.Y("mean(tuition_usd):Q", title="平均总学费（美元）"),
            color=alt.Color("country:N", legend=None),
            tooltip=[
                alt.Tooltip("country:N", title="国家/地区"),
                alt.Tooltip("mean(tuition_usd):Q", title="平均学费", format=",.0f"),
            ],
        )
        .properties(height=320)
    )
    gpa_chart = (
        alt.Chart(program_data)
        .mark_bar()
        .encode(
            x=alt.X("country:N", title="国家/地区"),
            y=alt.Y("mean(min_gpa_100):Q", title="平均最低 GPA"),
            color=alt.Color("country:N", legend=None),
            tooltip=[
                alt.Tooltip("country:N", title="国家/地区"),
                alt.Tooltip("mean(min_gpa_100):Q", title="平均最低 GPA", format=".1f"),
            ],
        )
        .properties(height=320)
    )
    left_column, right_column = st.columns(2)
    with left_column:
        st.altair_chart(cost_chart, use_container_width=True)
    with right_column:
        st.altair_chart(gpa_chart, use_container_width=True)

    budget_table = results[["display_name", "tuition_usd", "fit_score", "recommendation"]].copy()
    budget_table["是否在预算内"] = budget_table["tuition_usd"] <= profile["budget_usd"]
    budget_table = budget_table.rename(
        columns={
            "display_name": "项目",
            "tuition_usd": "学费USD",
            "fit_score": "匹配指数",
            "recommendation": "建议",
        }
    )
    st.dataframe(budget_table, use_container_width=True, hide_index=True)


def show_data_tab(program_data: pd.DataFrame) -> None:
    st.subheader("样例数据")
    st.caption("数据为课堂演示样例，不代表院校最新官方录取标准。真实申请前需以学校官网为准。")
    st.dataframe(program_data, use_container_width=True, hide_index=True)
    st.download_button(
        "下载样例数据",
        program_data.to_csv(index=False).encode("utf-8-sig"),
        file_name="masters_admission_sample.csv",
        mime="text/csv",
    )


def main() -> None:
    st.title("海外硕士录取标准分析模型")
    st.caption("围绕商科/数据类硕士申请，展示硬门槛、专业匹配和软实力三类因素的交互式分析。")

    program_data = load_program_data()
    profile, filtered_programs = build_sidebar(program_data)

    with st.expander("使用说明", expanded=False):
        st.write(
            "左侧输入申请人背景，页面会根据样例项目数据计算录取匹配指数，并输出保底、匹配、冲刺或暂不达标。"
            "本工具强调 Python 数据分析和可解释规则，不替代真实申请评估。"
        )

    if filtered_programs.empty:
        st.warning("当前国家/地区或方向筛选后没有项目，请调整左侧筛选项。")
        return

    results = calculate_match_results(filtered_programs, profile)
    show_metrics(results, profile)
    show_auto_advice(results)

    standards_tab, match_tab, model_tab, cost_tab, data_tab = st.tabs(
        ["项目门槛", "匹配结果", "模型解释", "费用对比", "样例数据"]
    )
    with standards_tab:
        show_standards_tab(filtered_programs)
    with match_tab:
        show_match_tab(results)
    with model_tab:
        show_model_tab(results)
    with cost_tab:
        show_cost_tab(filtered_programs, results, profile)
    with data_tab:
        show_data_tab(filtered_programs)


if __name__ == "__main__":
    main()
