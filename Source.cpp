#include "Scheduler.hpp"

void main() {
	Scheduler* scheduler = new Scheduler("Z Task name", "IwantSleep");
	scheduler->schedule();

	delete scheduler;
}