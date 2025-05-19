About This Repository

This repository contains a VBA-driven Excel tool that automates air balance calculations and psychrometric analysis based on .csv outputs from Trane Trace 3D Plus. Designed by a licensed mechanical engineer and GT MSCS candidate, this tool bridges the gap between load calculation software and real-world air handling unit (AHU) design and verification workflows.

🛠 Key Capabilities
	•	Automated Parsing of Trace 3D Plus .csv output files to extract space- and zone-level data
	•	Dynamic Air Balancing that tracks Supply, Return, Exhaust, and Outdoor Air for each space
	•	Pressurization Verification for all positively and negatively pressurized rooms, including transfer airflow tracking
	•	ASHRAE 170 Compliance mapping of desired vs. actual pressure relationships and airflow offsets
	•	Psychrometric Summaries with dewpoint, dry bulb, wet bulb, enthalpy, and RH% per AHU
	•	Airflow-Based Calculations like CFM/Ton, sensible and latent gains, space SHR, reheat BTUH, and more

📊 Outputs
	•	Comprehensive air balance tables (SA, RA, EA, OA, offset) by space
	•	Room-by-room pressure verification and directionality
	•	AHU-level summaries of psychrometric conditions and system loads
	•	Automatically color-coded flags for pressurization errors, mismatched design airflow, or imbalance

💡 Use Cases
	•	Engineers performing QA/QC on healthcare and mission-critical HVAC designs
	•	Commissioning agents verifying air balance integrity pre-occupancy
	•	Energy modelers and controls engineers aligning Trace 3D results with sequence-of-operation logic
	•	Project managers needing quick, automated space-level summaries for AHU control sequences

⚙ Technical Overview
	•	Built entirely in VBA for Excel
	•	Modular macros that:
	•	Import raw .csv outputs from Trace 3D Plus
	•	Parse and map each space to its AHU
	•	Calculate desired vs. actual airflow
	•	Summarize thermodynamic and control-relevant parameters
	•	Designed for use across healthcare, laboratory, data center, and other complex HVAC systems
 
